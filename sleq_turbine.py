"""
============================================================
SLEQ STREAMLINE CURVATURE METHOD — AXIAL TURBINE STAGE
1D Radial Analysis at Hub / Mean / Tip
============================================================

THEORY
------
The Simple Radial Equilibrium Equation (SREE) with streamline
curvature correction:

    dCm²/dr = 2*(Cθ/r)*(d(r*Cθ)/dr) / r
              - 2*dh0/dr
              + 2*T*ds/dr
              - Cm²*(dCm/dm)*(2/rc)    ← curvature term

For isentropic stator (ds=0, dh0=0) and design intent dh0/dr=0
(uniform work input), the SLEQ reduces to:

    dCm²/dr = -(2*Cθ/r) * d(Cθ)/dr + (2*Cθ²/r)    [SRE part]
              - Cm² * κ                              [curvature]

where κ = 2/rc  (rc = streamline radius of curvature)

METHOD
------
1. Mean-line solution (standard Euler / velocity triangle)
2. Assume a swirl distribution law: Cθ = f(r)
   - Free vortex:     r * Cθ = const  (most common)
   - Forced vortex:   Cθ = K*r
   - Solid body mix:  Cθ = A/r + B*r
3. Integrate the SLEQ ODE radially from mean to hub/tip
   using scipy.integrate.solve_ivp (RK45)
4. At each station: compute full velocity triangle,
   Mach numbers, blade angles, local loading, de Haller,
   Zweifel, choking check, h/D check

SIGN CONVENTION
---------------
- All angles measured from the axial direction
- Positive Cθ in direction of blade rotation
- Radial coordinate r from axis
- Blade speed U = Ω * r  (varies with radius)
"""

import numpy as np
from scipy.integrate import solve_ivp
from scipy.optimize import brentq
import sys

# ── tolerances ───────────────────────────────────────────────────────────────
TOL_ITER = 1e-6
MAX_ITER = 200

# ============================================================
# INPUT PARAMETERS  ← edit here
# ============================================================
class Inputs:
    # Boundary conditions
    P01     = 400_000.0   # Pa    total inlet pressure
    T01     = 1_200.0     # K     total inlet temperature
    P03     = 200_000.0   # Pa    stage exit total pressure
    mdot    = 10.0        # kg/s  mass flow

    # Gas
    gamma   = 1.333       # –
    Cp      = 1148.0      # J/kg·K

    # Stage design (mean-line)
    D_hub   = 0.250       # m     hub diameter
    alpha1  = 70.0        # deg   nozzle (stator) exit angle at mean
    Lambda  = 0.500       # –     stage reaction at mean
    psi     = 1.800       # –     loading coeff ΔH/U² at mean
    sc      = 0.850       # –     pitch-to-chord (both rows)
    lap     = 0.005       # m     rotor overhang = h_rotor - h_stator

    # Blade geometry
    chord_s = 0.060       # m     stator chord
    chord_r = 0.055       # m     rotor chord

    # Efficiencies
    eta_tt  = 0.890       # –     total-to-total isentropic
    eta_p   = 0.910       # –     polytropic

    # Swirl distribution law: 'free_vortex' | 'forced_vortex' | 'solid_body'
    swirl_law = 'free_vortex'

    # Streamline curvature at each blade row (radius of curvature [m])
    # None → curvature term is zero (= pure SRE)
    rc_stator = 0.15      # m     representative rc at stator TE
    rc_rotor  = 0.12      # m     representative rc at rotor TE

    # Number of integration steps (between mean and hub/tip)
    n_int = 200

INP = Inputs()

# ── derived gas constants ─────────────────────────────────────────────────────
R_gas  = INP.Cp * (INP.gamma - 1) / INP.gamma
gm1    = INP.gamma - 1
gp1    = INP.gamma + 1
gm1_g  = gm1 / INP.gamma

# ── trig helpers ─────────────────────────────────────────────────────────────
def sind(d): return np.sin(np.radians(d))
def cosd(d): return np.cos(np.radians(d))
def tand(d): return np.tan(np.radians(d))
def atan2d(y, x): return np.degrees(np.arctan2(y, x))
def atand(x): return np.degrees(np.arctan(x))

# ============================================================
# STEP 1 — MEAN-LINE SOLUTION
# ============================================================
def solve_meanline():
    """
    Euler turbomachinery + loading/reaction → blade speed U,
    axial velocity Cx, all mean-line angles & thermodynamic states.
    Returns a dict of mean-line quantities.
    """
    g   = INP.gamma
    Cp  = INP.Cp
    P01 = INP.P01
    T01 = INP.T01

    # Isentropic & actual stage work
    dh0s   = Cp * T01 * (1 - (INP.P03/P01)**gm1_g)
    dh0    = INP.eta_tt * dh0s          # J/kg  actual specific work
    T03    = T01 - dh0/Cp

    # Velocity triangle coefficients at mean
    Cu2_U  = (1 - INP.Lambda) + INP.psi/2   # Cu2/U
    Cu3_U  = (1 - INP.Lambda) - INP.psi/2   # Cu3/U
    phi    = Cu2_U / tand(INP.alpha1)        # φ = Cx/U
    C2_U   = np.sqrt(phi**2 + Cu2_U**2)     # C2/U

    # Blade speed from loading
    U      = np.sqrt(dh0 / INP.psi)         # m/s
    Cx     = U * phi                         # axial velocity (uniform radially)
    Omega  = 2*U / (INP.D_hub + 2*INP.D_hub)  # placeholder, corrected below

    # State 2 statics (isentropic stator)
    C2     = U * C2_U
    T2     = T01 - C2**2/(2*Cp)
    P2     = P01 * (T2/T01)**(g/gm1)
    rho2   = P2 / (R_gas * T2)
    Cu2    = U * Cu2_U
    Cu3    = U * Cu3_U

    # Annulus geometry (from continuity at state 2)
    A2     = INP.mdot / (rho2 * Cx)
    disc   = INP.D_hub**2 + 4*A2/np.pi
    h_s    = (-INP.D_hub + np.sqrt(disc)) / 2   # stator blade height
    D_mean = INP.D_hub + h_s
    D_tip  = INP.D_hub + 2*h_s
    h_r    = h_s + INP.lap                       # rotor blade height

    # Correct Omega
    r_mean = D_mean / 2
    Omega  = U / r_mean
    N_rpm  = Omega * 60 / (2*np.pi)

    # State 3 statics
    C3     = np.sqrt(Cx**2 + Cu3**2)
    T03    = T01 - dh0/Cp
    T3     = T03 - C3**2/(2*Cp)

    # Spouting velocity
    C0     = np.sqrt(2*dh0s)

    return dict(
        dh0s=dh0s, dh0=dh0, T03=T03, C0=C0,
        U=U, Cx=Cx, phi=phi, Omega=Omega, N_rpm=N_rpm,
        Cu2=Cu2, Cu3=Cu3, C2=C2, C3=C3,
        T2=T2, P2=P2, rho2=rho2,
        T3=T3,
        D_hub=INP.D_hub, D_mean=D_mean, D_tip=D_tip,
        r_hub=INP.D_hub/2, r_mean=D_mean/2, r_tip=D_tip/2,
        h_s=h_s, h_r=h_r, A2=A2,
        Cu2_U=Cu2_U, Cu3_U=Cu3_U,
    )


# ============================================================
# STEP 2 — SWIRL DISTRIBUTION LAW
# ============================================================
def swirl_distribution(r, r_mean, Cu_mean, law, row='stator'):
    """
    Return Cθ(r) given the law and mean value.

    Free vortex:  r·Cθ = const  → Cθ = Cu_mean * r_mean / r
    Forced vortex: Cθ = Cu_mean * r / r_mean
    Solid body:   Cθ = A/r + B*r  (A,B fitted to mean + hub condition)
    """
    if law == 'free_vortex':
        return Cu_mean * r_mean / r
    elif law == 'forced_vortex':
        return Cu_mean * r / r_mean
    elif law == 'solid_body':
        # Equal mix: A = B*r_mean² solved from Cu_mean
        # Cθ(r_mean) = A/r_mean + B*r_mean = Cu_mean
        # choose A = B * r_mean²  → 2B*r_mean = Cu_mean → B = Cu_mean/(2*r_mean)
        B = Cu_mean / (2 * r_mean)
        A = B * r_mean**2
        return A/r + B*r
    else:
        raise ValueError(f"Unknown swirl law: {law}")


def d_swirl_dr(r, r_mean, Cu_mean, law):
    """
    Analytical derivative dCθ/dr.
    """
    if law == 'free_vortex':
        return -Cu_mean * r_mean / r**2
    elif law == 'forced_vortex':
        return Cu_mean / r_mean
    elif law == 'solid_body':
        B = Cu_mean / (2 * r_mean)
        A = B * r_mean**2
        return -A/r**2 + B
    else:
        raise ValueError(f"Unknown swirl law: {law}")


# ============================================================
# STEP 3 — SLEQ ODE with Streamline Curvature
# ============================================================
def build_sleq_rhs(r_mean, Cu_mean, Cx_mean, Omega, law, rc, T0, P0, row='stator'):
    """
    Returns a callable f(r, y) for scipy.integrate.solve_ivp
    where y = [Cm²]  (meridional velocity squared).

    SLEQ with curvature:
        dCm²/dr = (2*Cθ/r)*d(r·Cθ)/dr / r
                  - Cm² * (2/rc) * sign(r-r_mean)
    
    Note: d(r·Cθ)/dr = Cθ + r*dCθ/dr
    
    For uniform work (dh0/dr=0) and isentropic (ds/dr=0):
        dCm²/dr = 2*Cθ*(dCθ/dr) + 2*Cθ²/r - Cm²*(2/rc)*direction
    """
    def rhs(r, y):
        Cm2 = float(y[0])
        if Cm2 < 0:
            Cm2 = 1.0  # prevent sqrt of negative during stiff step

        Cu    = swirl_distribution(r, r_mean, Cu_mean, law)
        dCu   = d_swirl_dr(r, r_mean, Cu_mean, law)

        # Swirl source term (always present)
        sre_term = 2*Cu*dCu + 2*Cu**2/r

        # Streamline curvature correction
        # Direction: integration outward (hub→tip) has opposite sign to inward
        if rc is not None and rc > 0:
            curv_term = Cm2 * (2.0/rc)
        else:
            curv_term = 0.0

        dCm2_dr = sre_term - curv_term
        return [dCm2_dr]

    return rhs


def integrate_sleq(r_start, r_end, Cm2_start, r_mean, Cu_mean, Omega,
                   law, rc, T01, P01, n_pts=200, row='stator'):
    """
    Integrate the SLEQ ODE from r_start to r_end.
    Returns arrays of (r, Cm, Cθ) along the span.
    """
    r_span  = (r_start, r_end)
    y0      = [Cm2_start]
    r_eval  = np.linspace(r_start, r_end, n_pts)

    rhs = build_sleq_rhs(r_mean, Cu_mean, Cm2_start**0.5 * 0,
                         Omega, law, rc, T01, P01, row)
    # Rebuild with correct Cu_mean
    rhs = build_sleq_rhs(r_mean, Cu_mean, 0, Omega, law, rc, T01, P01, row)

    sol = solve_ivp(rhs, r_span, y0, t_eval=r_eval,
                    method='RK45', rtol=1e-8, atol=1e-10,
                    dense_output=True)

    r_arr   = sol.t
    Cm2_arr = sol.y[0]
    Cm_arr  = np.sqrt(np.maximum(Cm2_arr, 0.01))
    Cu_arr  = np.array([swirl_distribution(r, r_mean, Cu_mean, law) for r in r_arr])

    return r_arr, Cm_arr, Cu_arr


# ============================================================
# STEP 4 — LOCAL STATION ANALYSIS
# ============================================================
def station_analysis(r, Cm, Cu_abs, Omega, T0, P0, Cp, gamma,
                     chord, sc_val, Cu_abs_upstream=None, Cm_upstream=None,
                     is_rotor=False, row_label='stator'):
    """
    Full velocity triangle + thermodynamic + geometry analysis at radius r.

    Parameters
    ----------
    r            : local radius [m]
    Cm           : local meridional (axial) velocity [m/s]
    Cu_abs       : local absolute tangential velocity [m/s]  (this row EXIT)
    Omega        : angular velocity [rad/s]
    T0           : upstream stagnation temperature [K]
    P0           : upstream stagnation pressure [Pa]
    chord        : blade chord [m]
    sc_val       : pitch-to-chord s/c
    Cu_abs_upstream : absolute Cθ at this row INLET (for de Haller)
    Cm_upstream  : Cm at inlet of this row (for de Haller)
    is_rotor     : True for rotor row
    """
    U  = Omega * r
    g  = gamma
    gm1_loc = g - 1

    # ── Absolute velocity at exit ──────────────────────────────────────────
    C      = np.sqrt(Cm**2 + Cu_abs**2)
    alpha  = atand(Cu_abs / Cm) if Cm > 0 else 90.0   # absolute exit angle

    # ── Static conditions (isentropic from upstream T0, P0) ──────────────
    T_stat = T0 - C**2 / (2*Cp)
    if T_stat <= 0:
        T_stat = 1.0
    P_stat = P0 * (T_stat/T0)**(g/gm1_loc)
    rho    = P_stat / (R_gas * T_stat)
    a_snd  = np.sqrt(g * R_gas * T_stat)
    Ma_abs = C / a_snd

    # ── Relative velocity at exit ─────────────────────────────────────────
    Wu     = Cu_abs - U          # relative tangential component
    W      = np.sqrt(Cm**2 + Wu**2)
    beta   = atand(Wu / Cm) if Cm > 0 else 90.0   # relative exit angle
    Ma_rel = W / a_snd

    # ── Local loading & reaction (for rotor) ─────────────────────────────
    if is_rotor and Cu_abs_upstream is not None:
        Cu_in  = Cu_abs_upstream   # absolute Cθ at rotor inlet
        Wu_in  = Cu_in - U         # relative Wu at rotor inlet
        W_in   = np.sqrt((Cm_upstream or Cm)**2 + Wu_in**2)
        beta_in = atand(Wu_in / (Cm_upstream or Cm))

        dH_local = U * (Cu_in - Cu_abs)   # Euler work at this radius
        psi_local = dH_local / U**2 if U > 0 else 0.0

        # de Haller number for rotor: W_exit / W_inlet
        de_haller = W / W_in if W_in > 0 else 1.0

        # Reaction at this radius
        # Λ = 1 - (Cu_in + Cu_abs)/(2U)
        Lambda_local = 1 - (Cu_in + Cu_abs)/(2*U) if U > 0 else INP.Lambda

        # Zweifel for rotor
        s_local  = sc_val * chord
        Zw = (2 * s_local/chord
              * np.cos(np.radians(beta))**2
              * (np.tan(np.radians(beta_in)) + abs(np.tan(np.radians(beta)))))
    else:
        beta_in    = atand((Cu_abs_upstream - U)/(Cm_upstream or Cm)) if (Cu_abs_upstream is not None and Cm_upstream is not None) else 0.0
        W_in       = np.sqrt((Cm_upstream or Cm)**2 + ((Cu_abs_upstream or Cu_abs) - U)**2) if Cu_abs_upstream is not None else W
        Wu_in      = (Cu_abs_upstream or Cu_abs) - U if Cu_abs_upstream is not None else Wu
        dH_local   = 0.0
        psi_local  = 0.0
        de_haller  = Cm / (Cm_upstream or Cm) if (Cm_upstream and not is_rotor) else None
        Lambda_local = INP.Lambda
        s_local    = sc_val * chord
        # Zweifel for stator (alpha_in ≈ 0 axial)
        alpha_in   = atand((Cu_abs_upstream or 0) / (Cm_upstream or Cm)) if (Cu_abs_upstream is not None and Cm_upstream is not None) else 0.0
        Zw = (2 * s_local/chord
              * np.cos(np.radians(alpha))**2
              * (abs(np.tan(np.radians(alpha_in))) + abs(np.tan(np.radians(alpha)))))

    # ── Blade count from local pitch at this radius ───────────────────────
    s_geom   = sc_val * chord
    Z_local  = round(2*np.pi*r / s_geom)
    s_actual = 2*np.pi*r / Z_local if Z_local > 0 else s_geom
    sc_actual = s_actual / chord

    # ── PASS / FAIL checks ───────────────────────────────────────────────
    choke_abs   = Ma_abs >= 1.0
    choke_rel   = Ma_rel >= 1.0
    dh_pass     = (de_haller is None) or (de_haller >= 0.72)   # de Haller ≥ 0.72
    zw_lo, zw_hi = (0.8, 1.0) if is_rotor else (0.7, 1.1)
    zw_pass     = zw_lo <= Zw <= zw_hi

    return dict(
        r=r, U=U, Cm=Cm,
        C=C, Cu_abs=Cu_abs, alpha=alpha, Ma_abs=Ma_abs,
        W=W, Wu=Wu, beta=beta, Ma_rel=Ma_rel,
        T_stat=T_stat, P_stat=P_stat, rho=rho,
        dH_local=dH_local, psi_local=psi_local,
        de_haller=de_haller, Lambda_local=Lambda_local,
        beta_in=beta_in if is_rotor else None,
        alpha_in=alpha_in if not is_rotor else None,
        Zw=Zw, chord=chord, sc_actual=sc_actual,
        Z_local=Z_local,
        choke_abs=choke_abs, choke_rel=choke_rel,
        dh_pass=dh_pass, zw_pass=zw_pass,
        is_rotor=is_rotor,
    )


# ============================================================
# STEP 5 — MASS FLOW CORRECTION (continuity check)
# ============================================================
def mass_flow_integral(r_arr, Cm_arr, rho_arr):
    """
    Numerically integrate ṁ = ∫ ρ·Cm · 2π·r dr  over span.
    Uses trapezoidal rule.
    """
    integrand = rho_arr * Cm_arr * 2*np.pi * r_arr
    return np.trapezoid(integrand, r_arr)


# ============================================================
# MAIN SOLVER
# ============================================================
def run():
    print("=" * 70)
    print("  SLEQ STREAMLINE CURVATURE SOLVER — AXIAL TURBINE STAGE")
    print("  3 Stations: Hub | Mean | Tip")
    print("=" * 70)

    # ── Step 1: Mean-line ────────────────────────────────────────────────
    ml = solve_meanline()
    r_hub  = ml['r_hub']
    r_mean = ml['r_mean']
    r_tip  = ml['r_tip']
    Omega  = ml['Omega']
    Cx     = ml['Cx']          # mean axial velocity (starting point)
    Cu2_m  = ml['Cu2']         # mean absolute Cθ at stator exit
    Cu3_m  = ml['Cu3']         # mean absolute Cθ at rotor exit
    dh0    = ml['dh0']

    print(f"\n{'─'*70}")
    print("  MEAN LINE SOLUTION")
    print(f"{'─'*70}")
    print(f"  Blade speed U_mean    = {ml['U']:10.3f} m/s")
    print(f"  Axial velocity Cx     = {Cx:10.3f} m/s")
    print(f"  Flow coefficient φ    = {ml['phi']:10.4f}")
    print(f"  Rotational speed N    = {ml['N_rpm']:10.1f} rpm")
    print(f"  D_hub / mean / tip    = {2*r_hub*1000:.1f} / {2*r_mean*1000:.1f} / {2*r_tip*1000:.1f} mm")
    print(f"  Stator height h_s     = {ml['h_s']*1000:.2f} mm")
    print(f"  Rotor height h_r      = {ml['h_r']*1000:.2f} mm")
    print(f"  Actual work ΔH        = {dh0/1000:.3f} kJ/kg")
    print(f"  Swirl law             = {INP.swirl_law}")

    # ── Step 2 & 3: SLEQ integration — STATOR EXIT ───────────────────────
    # Integrate inward (mean→hub) and outward (mean→tip)
    n   = INP.n_int
    law = INP.swirl_law

    # STATOR EXIT (state 2): Cu varies with r, Cm varies from SLEQ
    r_in2, Cm2_in, Cu2_in = integrate_sleq(
        r_mean, r_hub, Cx**2,
        r_mean, Cu2_m, Omega, law, INP.rc_stator,
        INP.T01, INP.P01, n_pts=n, row='stator')

    r_out2, Cm2_out, Cu2_out = integrate_sleq(
        r_mean, r_tip, Cx**2,
        r_mean, Cu2_m, Omega, law, INP.rc_stator,
        INP.T01, INP.P01, n_pts=n, row='stator')

    # Interpolate at hub, mean, tip
    Cm2_hub  = float(np.interp(r_hub,  r_in2,  Cm2_in))
    Cu2_hub  = float(np.interp(r_hub,  r_in2,  Cu2_in))
    Cm2_mean = Cx  # by construction
    Cu2_mean = Cu2_m
    Cm2_tip  = float(np.interp(r_tip,  r_out2, Cm2_out))
    Cu2_tip  = float(np.interp(r_tip,  r_out2, Cu2_out))

    # ROTOR EXIT (state 3): same procedure
    r_in3, Cm3_in, Cu3_in = integrate_sleq(
        r_mean, r_hub, Cx**2,
        r_mean, Cu3_m, Omega, law, INP.rc_rotor,
        INP.T01, INP.P01, n_pts=n, row='rotor')

    r_out3, Cm3_out, Cu3_out = integrate_sleq(
        r_mean, r_tip, Cx**2,
        r_mean, Cu3_m, Omega, law, INP.rc_rotor,
        INP.T01, INP.P01, n_pts=n, row='rotor')

    Cm3_hub  = float(np.interp(r_hub, r_in3, Cm3_in))
    Cu3_hub  = float(np.interp(r_hub, r_in3, Cu3_in))
    Cm3_mean = Cx
    Cu3_mean = Cu3_m
    Cm3_tip  = float(np.interp(r_tip, r_out3, Cm3_out))
    Cu3_tip  = float(np.interp(r_tip, r_out3, Cu3_out))

    # ── Step 4: Station analysis ──────────────────────────────────────────
    stations_r = [r_hub, r_mean, r_tip]
    labels_r   = ['HUB', 'MEAN', 'TIP']

    Cm2_arr  = [Cm2_hub,  Cm2_mean,  Cm2_tip]
    Cu2_arr  = [Cu2_hub,  Cu2_mean,  Cu2_tip]
    Cm3_arr  = [Cm3_hub,  Cm3_mean,  Cm3_tip]
    Cu3_arr  = [Cu3_hub,  Cu3_mean,  Cu3_tip]

    # Stator inlet: assume axial (Cu=0, Cm≈Cx for all stations)
    # (small inlet Mach, uniform upstream)
    Cu1_arr  = [0.0, 0.0, 0.0]
    Cm1_arr  = [Cx, Cx, Cx]

    stator_results = []
    rotor_results  = []

    for i, (r, lbl) in enumerate(zip(stations_r, labels_r)):
        # ── STATOR analysis ──────────────────────────────────────────────
        sr = station_analysis(
            r       = r,
            Cm      = Cm2_arr[i],
            Cu_abs  = Cu2_arr[i],
            Omega   = Omega,
            T0      = INP.T01,
            P0      = INP.P01,
            Cp      = INP.Cp,
            gamma   = INP.gamma,
            chord   = INP.chord_s,
            sc_val  = INP.sc,
            Cu_abs_upstream = Cu1_arr[i],
            Cm_upstream     = Cm1_arr[i],
            is_rotor = False,
            row_label = f"Stator {lbl}",
        )
        stator_results.append((lbl, sr))

        # ── ROTOR analysis ───────────────────────────────────────────────
        # T0 for rotor = T03 (exit stagnation)
        # P0 for rotor: use isentropic approximation from T ratio
        T0r  = ml['T03']
        P0r  = INP.P03   # stage exit total pressure (approximate; full loss model omitted)

        rr = station_analysis(
            r       = r,
            Cm      = Cm3_arr[i],
            Cu_abs  = Cu3_arr[i],
            Omega   = Omega,
            T0      = T0r,
            P0      = P0r,
            Cp      = INP.Cp,
            gamma   = INP.gamma,
            chord   = INP.chord_r,
            sc_val  = INP.sc,
            Cu_abs_upstream = Cu2_arr[i],   # rotor inlet = stator exit
            Cm_upstream     = Cm2_arr[i],
            is_rotor = True,
            row_label = f"Rotor {lbl}",
        )
        rotor_results.append((lbl, rr))

    # ── Mass flow check ──────────────────────────────────────────────────
    # Build density arrays for continuity check at state 2
    rho2_arr = []
    for r, Cm, Cu in zip(stations_r, Cm2_arr, Cu2_arr):
        C   = np.sqrt(Cm**2 + Cu**2)
        T_s = INP.T01 - C**2/(2*INP.Cp)
        P_s = INP.P01*(T_s/INP.T01)**(INP.gamma/gm1)
        rho2_arr.append(P_s/(R_gas*T_s))

    # Trapezoidal mass flow over 3 points
    r_arr_3   = np.array(stations_r)
    rho2_3    = np.array(rho2_arr)
    Cm2_3     = np.array(Cm2_arr)
    mdot_check = np.trapezoid(rho2_3 * Cm2_3 * 2*np.pi * r_arr_3, r_arr_3)
    mdot_err   = abs(mdot_check - INP.mdot) / INP.mdot * 100

    # ── Print results ─────────────────────────────────────────────────────
    PASS = "✓ PASS"
    FAIL = "✗ FAIL"

    def pfmt(cond): return PASS if cond else FAIL

    row_sep = "─" * 70

    for row_name, results_list in [("STATOR (Row 1) — EXIT Conditions (State 2)",  stator_results),
                                    ("ROTOR  (Row 2) — EXIT Conditions (State 3)", rotor_results)]:
        print(f"\n{'='*70}")
        print(f"  {row_name}")
        print(f"{'='*70}")

        is_rotor = (row_name.startswith("ROTOR"))

        for lbl, d in results_list:
            r_mm  = d['r'] * 1000
            h_r_mm = d['r'] / r_mean   # radius ratio

            print(f"\n  ┌{'─'*65}┐")
            print(f"  │  STATION: {lbl:6s}   r = {r_mm:.1f} mm   r/r_mean = {h_r_mm:.4f}          │")
            print(f"  ├{'─'*65}┤")
            print(f"  │  VELOCITY TRIANGLE                                              │")
            print(f"  │  Blade Speed      U      = {d['U']:8.3f} m/s                        │")
            print(f"  │  Meridional       Cm     = {d['Cm']:8.3f} m/s                        │")
            print(f"  │  Absolute tang.   Cθ     = {d['Cu_abs']:8.3f} m/s                        │")
            print(f"  │  Absolute vel.    C      = {d['C']:8.3f} m/s                        │")
            if is_rotor:
                print(f"  │  Relative tang.   Wθ     = {d['Wu']:8.3f} m/s                        │")
                print(f"  │  Relative vel.    W      = {d['W']:8.3f} m/s                        │")
            print(f"  ├{'─'*65}┤")
            print(f"  │  ANGLES                                                         │")
            if is_rotor:
                print(f"  │  Rotor inlet β₂  = {d['beta_in'] if d['beta_in'] is not None else 0.0:8.3f} °  (from axial)                 │")
                print(f"  │  Rotor exit  β₃  = {d['beta']:8.3f} °  (from axial)                 │")
            else:
                print(f"  │  Stator inlet α₀ = {d['alpha_in'] if d['alpha_in'] is not None else 0.0:8.3f} °  (axial entry)                │")
                print(f"  │  Stator exit  α₁ = {d['alpha']:8.3f} °  (from axial)                 │")
            print(f"  ├{'─'*65}┤")
            print(f"  │  MACH NUMBERS                                                   │")
            print(f"  │  Ma_abs  = {d['Ma_abs']:.4f}   Ma_rel = {d['Ma_rel']:.4f}                           │")
            print(f"  ├{'─'*65}┤")
            print(f"  │  THERMODYNAMICS                                                 │")
            print(f"  │  T_static = {d['T_stat']:7.2f} K   P_static = {d['P_stat']/1000:7.2f} kPa             │")
            print(f"  │  Density  = {d['rho']:7.4f} kg/m³                                    │")
            print(f"  ├{'─'*65}┤")
            print(f"  │  LOADING & PERFORMANCE                                          │")
            if is_rotor:
                print(f"  │  Local work ΔH   = {d['dH_local']/1000:7.3f} kJ/kg                             │")
                print(f"  │  Loading ψ_local = {d['psi_local']:7.4f}                                        │")
                print(f"  │  Reaction Λ_loc  = {d['Lambda_local']:7.4f}                                        │")
            dh_str = f"{d['de_haller']:.4f}" if d['de_haller'] is not None else "  N/A "
            print(f"  │  de Haller W₃/W₂ = {dh_str:6s}   Zweifel Zw = {d['Zw']:.4f}                  │")
            print(f"  ├{'─'*65}┤")
            print(f"  │  BLADE GEOMETRY AT THIS RADIUS                                  │")
            print(f"  │  Chord c        = {d['chord']*1000:7.2f} mm                                    │")
            print(f"  │  Pitch s (act.) = {d['sc_actual']*d['chord']*1000:7.2f} mm   s/c = {d['sc_actual']:.4f}            │")
            print(f"  │  Blade count Z  = {d['Z_local']:5d}                                          │")
            stag = abs(d['alpha'] if not is_rotor else (d['beta_in'] or 0) + d['beta']) / 2
            print(f"  │  Stagger angle  ≈ {stag:7.2f} °  (mid-camber approx)              │")
            print(f"  ├{'─'*65}┤")
            print(f"  │  PASS / FAIL CRITERIA                                           │")
            print(f"  │  Choke (abs)   Ma<1.0      → {pfmt(not d['choke_abs']):8s}  [{d['Ma_abs']:.4f}]          │")
            print(f"  │  Choke (rel)   Ma_rel<1.0  → {pfmt(not d['choke_rel']):8s}  [{d['Ma_rel']:.4f}]          │")
            dh_val = f"{d['de_haller']:.4f}" if d['de_haller'] is not None else "  N/A"
            print(f"  │  de Haller     W₃/W₂≥0.72 → {pfmt(d['dh_pass']):8s}  [{dh_val}]         │")
            print(f"  │  Zweifel Zw    0.7–1.1     → {pfmt(d['zw_pass']):8s}  [{d['Zw']:.4f}]          │")
            print(f"  └{'─'*65}┘")

    # ── Radial profiles summary table ─────────────────────────────────────
    print(f"\n{'='*70}")
    print("  RADIAL PROFILE SUMMARY — Stator Exit (State 2)")
    print(f"{'='*70}")
    hdr = f"  {'Stn':5s} {'r[mm]':7s} {'U[m/s]':8s} {'Cm[m/s]':8s} {'Cθ[m/s]':8s} {'α[°]':7s} {'Ma_abs':7s} {'Ma_rel':7s} {'Zw':6s}"
    print(hdr)
    print(f"  {'-'*68}")
    for lbl, d in stator_results:
        print(f"  {lbl:5s} {d['r']*1000:7.2f} {d['U']:8.3f} {d['Cm']:8.3f} {d['Cu_abs']:8.3f} "
              f"{d['alpha']:7.2f} {d['Ma_abs']:7.4f} {d['Ma_rel']:7.4f} {d['Zw']:6.4f}")

    print(f"\n{'='*70}")
    print("  RADIAL PROFILE SUMMARY — Rotor Exit (State 3)")
    print(f"{'='*70}")
    hdr2 = f"  {'Stn':5s} {'r[mm]':7s} {'U[m/s]':8s} {'Cm[m/s]':8s} {'Cθ[m/s]':8s} {'β[°]':7s} {'ΔH[kJ]':7s} {'ψ_loc':7s} {'Λ_loc':6s} {'dHall':6s}"
    print(hdr2)
    print(f"  {'-'*68}")
    for lbl, d in rotor_results:
        dh_str = f"{d['de_haller']:.4f}" if d['de_haller'] is not None else "  N/A"
        print(f"  {lbl:5s} {d['r']*1000:7.2f} {d['U']:8.3f} {d['Cm']:8.3f} {d['Cu_abs']:8.3f} "
              f"{d['beta']:7.2f} {d['dH_local']/1000:7.3f} {d['psi_local']:7.4f} {d['Lambda_local']:6.4f} {dh_str:6s}")

    # ── Continuity check ──────────────────────────────────────────────────
    print(f"\n{'='*70}")
    print("  CONTINUITY & GLOBAL CHECKS")
    print(f"{'='*70}")
    print(f"  Mass flow input  ṁ_in   = {INP.mdot:.4f} kg/s")
    print(f"  Mass flow check  ṁ_calc = {mdot_check:.4f} kg/s  (3-point trapz at state 2)")
    print(f"  Mass flow error          = {mdot_err:.3f} %  {'✓ OK' if mdot_err<5 else '✗ CHECK MESH'}")
    print(f"  Swirl law                = {INP.swirl_law}")
    print(f"  rc_stator                = {INP.rc_stator} m")
    print(f"  rc_rotor                 = {INP.rc_rotor} m")

    # ── Swirl distribution tables ─────────────────────────────────────────
    print(f"\n{'='*70}")
    print("  SWIRL DISTRIBUTION — Stator Exit (full span, 10 pts)")
    print(f"{'='*70}")
    r_full = np.linspace(r_hub, r_tip, 10)
    print(f"  {'r[mm]':7s} {'r/r_mean':9s} {'Cθ_SRE[m/s]':12s} {'ΔCm[m/s]':10s}")
    print(f"  {'-'*44}")
    Cm2_full_in  = np.interp(r_full[:len(r_full)//2+1], r_in2[::-1], Cm2_in[::-1])
    Cm2_full_out = np.interp(r_full[len(r_full)//2:],   r_out2,      Cm2_out)
    Cm2_full = np.concatenate([Cm2_full_in[:-1], Cm2_full_out])
    if len(Cm2_full) < len(r_full):
        Cm2_full = np.append(Cm2_full, Cm2_full[-1])
    for r_i, Cm_i in zip(r_full, Cm2_full[:len(r_full)]):
        Cu_i   = swirl_distribution(r_i, r_mean, Cu2_m, law)
        dCm    = Cm_i - Cx
        print(f"  {r_i*1000:7.2f}  {r_i/r_mean:9.4f}  {Cu_i:12.4f}  {dCm:10.4f}")

    print(f"\n{'─'*70}")
    print("  SOLVER COMPLETE")
    print(f"{'─'*70}\n")

    return stator_results, rotor_results, ml


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    stator_res, rotor_res, ml = run()
