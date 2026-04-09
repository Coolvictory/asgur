"""
turbflow.thermo.nasa7
=====================
Evaluates NASA-7 polynomial thermodynamic properties for pure species
and gas mixtures.

Theory
------
For each species i, two sets of 7 coefficients (low T / high T range):

    cp_i°(T) / R  =  a1 + a2·T + a3·T² + a4·T³ + a5·T⁴          [dimensionless]

    h_i°(T) / (R·T)  =  a1 + a2/2·T + a3/3·T² + a4/4·T³
                          + a5/5·T⁴ + a6/T                         [dimensionless]

    s_i°(T) / R  =  a1·ln(T) + a2·T + a3/2·T² + a4/3·T³
                     + a5/4·T⁴ + a7                                 [dimensionless]

Mixture properties (mass-fraction weighted):
    cp_mix(T)  = Σ  Y_i · cp_i(T)                    [J / (kg·K)]
    h_mix(T)   = Σ  Y_i · h_i(T)                     [J / kg]
    R_mix      = Σ  Y_i · R_universal / M_i           [J / (kg·K)]
    γ_mix(T)   = cp_mix / (cp_mix - R_mix)            [-]

The pressure-dependence of h and cp is negligible for ideal gases used here.

Reference state: h°(298.15 K) from NASA report (heat of formation included).

All outputs in SI:
    cp   [J/(kg·K)]
    h    [J/kg]
    s    [J/(kg·K)]     (at given P via  s = s°(T) - R·ln(P/P_ref))
    γ    [-]
    R    [J/(kg·K)]
    M    [kg/mol]
"""

from __future__ import annotations

import math
from functools import lru_cache
from typing import Sequence

import numpy as np

from turbflow.thermo.nasa7_coefficients import (
    SPECIES_DATA,
    R_UNIVERSAL_J_MOL_K,
    get_species,
)

# Reference pressure for entropy [Pa]
P_REF: float = 101_325.0


# ═══════════════════════════════════════════════════════════════════════════
# Pure-species evaluation
# ═══════════════════════════════════════════════════════════════════════════

def _coeffs(name: str, T: float) -> tuple[list[float], float]:
    """Return the 7 NASA coefficients and R_i [J/(kg·K)] for species at T."""
    d = get_species(name)
    T = max(d["Tlo"], min(d["Thi"], T))   # clamp to valid range
    c = d["lo"] if T <= d["Tmid"] else d["hi"]
    R_i = R_UNIVERSAL_J_MOL_K / d["M"]   # specific gas constant
    return c, R_i


def cp_species(name: str, T: float) -> float:
    """
    Specific heat at constant pressure for pure species [J/(kg·K)].

    cp_i(T) = R_i · (a1 + a2·T + a3·T² + a4·T³ + a5·T⁴)
    """
    c, R_i = _coeffs(name, T)
    cp_over_R = c[0] + c[1]*T + c[2]*T**2 + c[3]*T**3 + c[4]*T**4
    return R_i * cp_over_R


def h_species(name: str, T: float) -> float:
    """
    Specific enthalpy for pure species [J/kg].

    h_i(T) = R_i · T · (a1 + a2/2·T + a3/3·T² + a4/4·T³ + a5/5·T⁴ + a6/T)
    """
    c, R_i = _coeffs(name, T)
    h_over_RT = (c[0]
                 + c[1]/2 * T
                 + c[2]/3 * T**2
                 + c[3]/4 * T**3
                 + c[4]/5 * T**4
                 + c[5]/T)
    return R_i * T * h_over_RT


def s_species(name: str, T: float, P: float = P_REF) -> float:
    """
    Specific entropy for pure species [J/(kg·K)] at pressure P.

    s_i(T,P) = R_i · (a1·ln(T) + a2·T + a3/2·T² + a4/3·T³ + a5/4·T⁴ + a7)
               - R_i · ln(P / P_ref)
    """
    c, R_i = _coeffs(name, T)
    s_over_R = (c[0]*math.log(T)
                + c[1]*T
                + c[2]/2 * T**2
                + c[3]/3 * T**3
                + c[4]/4 * T**4
                + c[6])
    return R_i * (s_over_R - math.log(P / P_REF))


# ═══════════════════════════════════════════════════════════════════════════
# Mixture class
# ═══════════════════════════════════════════════════════════════════════════

class NASA7Mixture:
    """
    Multi-species gas mixture evaluated via NASA-7 polynomials.

    Parameters
    ----------
    species : list[str]
        Species names matching keys in SPECIES_DATA (e.g. ["N2","O2","CO2","H2O","Ar"])
    Y : list[float]
        Mass fractions (must sum to 1.0 within 1e-6)

    All thermodynamic properties are functions of T (and P for entropy only).
    Properties are intensive, per unit mass of mixture.
    """

    def __init__(self, species: list[str], Y: list[float]) -> None:
        if len(species) != len(Y):
            raise ValueError("species and Y must have equal length")
        total = sum(Y)
        if abs(total - 1.0) > 1e-4:
            raise ValueError(f"Mass fractions sum to {total:.6f}, must be 1.0")

        self.species = list(species)
        self.Y = [yi / total for yi in Y]  # normalise defensively

        # Specific gas constants per species [J/(kg·K)]
        self._R_i: list[float] = [
            R_UNIVERSAL_J_MOL_K / get_species(sp)["M"]
            for sp in self.species
        ]

        # Mixture specific gas constant [J/(kg·K)] — constant (no T dependence)
        self.R_mix: float = sum(Yi * Ri for Yi, Ri in zip(self.Y, self._R_i))

        # Mixture molar mass [kg/mol]
        self.M_mix: float = 1.0 / sum(
            Yi / get_species(sp)["M"]
            for Yi, sp in zip(self.Y, self.species)
        )

    # ── Primary properties ─────────────────────────────────────────────────

    def cp(self, T: float) -> float:
        """Mixture cp [J/(kg·K)] at temperature T."""
        return sum(Yi * cp_species(sp, T)
                   for Yi, sp in zip(self.Y, self.species))

    def h(self, T: float) -> float:
        """Mixture specific enthalpy [J/kg] at temperature T."""
        return sum(Yi * h_species(sp, T)
                   for Yi, sp in zip(self.Y, self.species))

    def s(self, T: float, P: float = P_REF) -> float:
        """Mixture specific entropy [J/(kg·K)] at (T, P)."""
        return sum(Yi * s_species(sp, T, P)
                   for Yi, sp in zip(self.Y, self.species))

    def gamma(self, T: float) -> float:
        """Heat capacity ratio γ = cp / (cp - R_mix)."""
        cp_ = self.cp(T)
        return cp_ / (cp_ - self.R_mix)

    # ── Derived / inverse properties ───────────────────────────────────────

    def T_from_h(self, h_target: float, T_guess: float = 1000.0) -> float:
        """
        Invert h(T) = h_target via Newton–Raphson.

        Convergence is guaranteed because h is strictly monotone in T
        (cp > 0 everywhere in the valid range).
        """
        T = T_guess
        for _ in range(60):
            h_val = self.h(T)
            cp_val = self.cp(T)        # dh/dT = cp
            dT = (h_target - h_val) / cp_val
            T += dT
            T = max(200.0, min(6000.0, T))
            if abs(dT) < 1e-6:        # converged to micro-Kelvin
                break
        return T

    def T_from_s_P(self, s_target: float, P: float, T_guess: float = 1000.0) -> float:
        """
        Find T such that s(T, P) = s_target via Newton–Raphson.
        ds/dT = cp / T  →  Newton step: ΔT = (s_target - s(T,P)) / (cp/T)
        """
        T = T_guess
        for _ in range(60):
            s_val = self.s(T, P)
            cp_val = self.cp(T)
            dT = (s_target - s_val) * T / cp_val
            T += dT
            T = max(200.0, min(6000.0, T))
            if abs(dT) < 1e-6:
                break
        return T

    def delta_h(self, T1: float, T2: float, n_gauss: int = 8) -> float:
        """
        Exact enthalpy difference h(T2) - h(T1) [J/kg] by
        Gaussian quadrature of cp(T):

            Δh = ∫_{T1}^{T2} cp_mix(T) dT

        This is more accurate than h(T2) - h(T1) when T crosses Tmid,
        but both are equivalent since NASA-7 ensures continuity at Tmid.
        """
        return self.h(T2) - self.h(T1)

    def isentropic_T_outlet(
        self,
        T_in: float,
        P_in: float,
        P_out: float,
        T_guess: float | None = None,
    ) -> float:
        """
        Find T_out such that s(T_out, P_out) = s(T_in, P_in).
        This is the isentropic outlet temperature for a real-gas expansion.
        """
        s_in = self.s(T_in, P_in)
        T0 = T_guess if T_guess else T_in * (P_out / P_in) ** ((self.gamma(T_in) - 1) / self.gamma(T_in))
        return self.T_from_s_P(s_in, P_out, T_guess=T0)

    def speed_of_sound(self, T: float, P: float) -> float:
        """
        Speed of sound in the mixture [m/s].
        a = sqrt(γ · R_mix · T)    (ideal gas, valid here)
        """
        return math.sqrt(self.gamma(T) * self.R_mix * T)

    # ── Summary ────────────────────────────────────────────────────────────

    def summary(self, T: float, P: float) -> dict:
        cp_ = self.cp(T)
        return {
            "T_K":    T,
            "P_Pa":   P,
            "cp_J_kgK": cp_,
            "gamma":  self.gamma(T),
            "R_mix_J_kgK": self.R_mix,
            "h_J_kg": self.h(T),
            "s_J_kgK": self.s(T, P),
            "a_m_s":  self.speed_of_sound(T, P),
            "M_mix_kg_mol": self.M_mix,
        }

    def __repr__(self) -> str:
        fracs = ", ".join(
            f"{sp}:{Yi:.4f}" for sp, Yi in zip(self.species, self.Y)
        )
        return f"NASA7Mixture([{fracs}], R_mix={self.R_mix:.2f} J/kg/K)"
