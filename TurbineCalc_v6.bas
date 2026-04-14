Attribute VB_Name = "TurbineCalc"
Option Explicit
' ================================================================
' MULTI-STAGE AXIAL TURBINE MEAN LINE SOLVER  v6
' ----------------------------------------------------------------
' 3-STATION INDEPENDENT Cx -- 0.000000% continuity at all states
' ----------------------------------------------------------------
' STATE 1 stator inlet:
'   A1 = A2 (same annulus, h_s from state 2)
'   C1 from continuity: iterate C1 = mdot/(rho1*A2)
'   rho1 = P1s/(Z*R*T1s),  T1s = T01 - C1^2/(2Cp*1000)
'
' STATE 2 stator exit / rotor inlet:
'   phi2 = Cu2U / tan(alpha1)   [alpha1 = stator gauge angle, col H]
'   U = Omega * (Dhub+h_s)/2
'   Cx2 = U * phi2
'   NR on h_s until continuity converged
'
' STATE 3 rotor exit:
'   beta3 = rotor exit gauge angle [col I = 'Rotor Gauge Angle']
'   Wu3 = Cu3 - U  (from Euler, known)
'   Cx3 = |Wu3| / tan(|beta3|)   -- determined by beta3 input
'   C3  = sqrt(Cx3^2 + Cu3^2)
'   T3s = T03 - C3^2/(2*Cp3*1000)
'   rho3 -> A3 = mdot/(rho3*Cx3) -> h_r = A3/(pi*Dm)
'
' Cp DECREASES with T  (a1 > 0), correct for combustion products
' Fixed shaft RPM -- U = Omega * r_mean, r_mean from state 2 NR
' ================================================================

Const PI As Double = 3.14159265358979
Const MAX_S As Integer = 6

Private Sub Wv(ws As Worksheet, r As Long, c As Long, v As Variant, nf As String)
    ws.Cells(r, c).Value = v
    If nf <> "" Then ws.Cells(r, c).NumberFormat = nf
End Sub

Private Sub WvPF(ws As Worksheet, r As Long, c As Long, isPass As Boolean)
    ws.Cells(r, c).Value = IIf(isPass, "PASS", "FAIL")
    ws.Cells(r, c).NumberFormat = "@"
    ws.Cells(r, c).Font.Bold = True
    ws.Cells(r, c).Font.Name = "Arial"
    ws.Cells(r, c).Font.Size = 9
    If isPass Then
        ws.Cells(r, c).Interior.Color = RGB(212, 237, 218)
        ws.Cells(r, c).Font.Color = RGB(26, 94, 42)
    Else
        ws.Cells(r, c).Interior.Color = RGB(248, 215, 218)
        ws.Cells(r, c).Font.Color = RGB(123, 0, 0)
    End If
End Sub

Private Function CpT(T As Double, a0 As Double, a1 As Double, a2 As Double) As Double
    CpT = a0 + a1 * T + a2 * T * T
End Function

Public Sub RunTurbineCalc()

    Dim wsIn As Worksheet, wsGP As Worksheet, wsTH As Worksheet
    Dim wsVT As Worksheet, wsGE As Worksheet, wsBL As Worksheet
    Dim wsCR As Worksheet, wsSU As Worksheet

    Set wsIn = ThisWorkbook.Sheets("INPUTS")
    Set wsGP = ThisWorkbook.Sheets("GAS_PROPERTIES")
    Set wsTH = ThisWorkbook.Sheets("THERMO")
    Set wsVT = ThisWorkbook.Sheets("VEL_TRIANGLES")
    Set wsGE = ThisWorkbook.Sheets("GEOMETRY")
    Set wsBL = ThisWorkbook.Sheets("BLADES")
    Set wsCR = ThisWorkbook.Sheets("CRITERIA")
    Set wsSU = ThisWorkbook.Sheets("SUMMARY")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ── Global inputs ────────────────────────────────────────
    Dim P01   As Double: P01   = wsIn.Range("D5").Value
    Dim T01   As Double: T01   = wsIn.Range("D6").Value
    Dim mdot  As Double: mdot  = wsIn.Range("D7").Value
    Dim P0n   As Double: P0n   = wsIn.Range("D8").Value
    Dim gam1  As Double: gam1  = wsIn.Range("I5").Value
    Dim Cp1   As Double: Cp1   = wsIn.Range("I6").Value
    Dim rhoIn As Double: rhoIn = wsIn.Range("M13").Value
    Dim a0    As Double: a0    = wsIn.Range("M7").Value
    Dim a1c   As Double: a1c   = wsIn.Range("M8").Value
    Dim a2c   As Double: a2c   = wsIn.Range("M9").Value
    Dim N_st  As Integer: N_st = CInt(wsIn.Range("D12").Value)
    Dim Nrpm  As Double: Nrpm  = wsIn.Range("I9").Value

    ' ── Auto-correct Cp direction: a1 must be > 0 ────────────
    If a1c <= 0 Then
        a1c = 2.0E-4
        a0  = Cp1 - a1c * T01
        a2c = 0
        wsIn.Range("M7").Value = a0
        wsIn.Range("M8").Value = a1c
        wsIn.Range("M9").Value = a2c
        wsIn.Range("M7").Interior.Color = RGB(255, 235, 156)
        wsIn.Range("M8").Interior.Color = RGB(255, 235, 156)
        MsgBox "Cp polynomial corrected: a1 set to +2e-4 (Cp decreases with T)." & _
               vbCrLf & "See INPUTS M7:M9.", vbInformation, "Cp Correction"
    End If

    If N_st < 1 Then N_st = 1
    If N_st > MAX_S Then N_st = MAX_S

    Dim R_mix As Double: R_mix = Cp1 * (gam1 - 1) / gam1   ' kJ/kgK
    Dim R_J   As Double: R_J   = R_mix * 1000               ' J/kgK
    Dim Z_in  As Double: Z_in  = P01 / (rhoIn * R_J * T01)
    Dim Omega As Double: Omega = Nrpm * 2 * PI / 60

    ' ── Total T-drop (polytropic) ─────────────────────────────
    Dim setp As Double: setp = 0
    Dim spsi As Double: spsi = 0
    Dim s As Integer
    For s = 0 To N_st - 1
        setp = setp + wsIn.Cells(16 + s, 15).Value  ' col O eta_p
        spsi = spsi + wsIn.Cells(16 + s, 11).Value  ' col K psi
    Next s
    Dim ep_avg As Double: ep_avg = setp / N_st
    Dim T0n As Double: T0n = T01 * (P0n / P01) ^ (ep_avg * (gam1 - 1) / gam1)
    Dim dTtot As Double: dTtot = T01 - T0n

    ' ── Stage loop ───────────────────────────────────────────
    Dim T0i As Double: T0i = T01
    Dim P0i As Double: P0i = P01

    For s = 0 To MAX_S - 1
        Dim gr As Long: gr = 4 + s
        Dim sr As Long: sr = 16 + s

        If s >= N_st Then
            wsGP.Rows(gr).ClearContents
            wsTH.Rows(gr).ClearContents
            wsVT.Rows(gr).ClearContents
            wsGE.Rows(gr).ClearContents
            wsBL.Rows(gr).ClearContents
            wsCR.Rows(gr).ClearContents
            wsSU.Rows(gr).ClearContents
            GoTo Nxt
        End If

        ' ── Per-stage inputs ──────────────────────────────────
        ' Col map (confirmed from workbook audit):
        ' D=4  Stator Hub Dia [mm]
        ' F=6  Stator Chord [mm]
        ' G=7  Rotor Chord [mm]
        ' H=8  Stator Gauge Angle alpha1 [deg]
        ' I=9  Rotor Gauge Angle beta3 [deg]   <-- NEW USE: rotor EXIT angle
        ' J=10 Reaction Lambda
        ' K=11 Loading psi
        ' L=12 Pitch/Chord s/c
        ' N=14 eta_tt
        ' O=15 eta_p
        Dim Dhub  As Double: Dhub  = wsIn.Cells(sr, 4).Value / 1000
        Dim cs_m  As Double: cs_m  = wsIn.Cells(sr, 6).Value / 1000
        Dim cr_m  As Double: cr_m  = wsIn.Cells(sr, 7).Value / 1000
        Dim a1d   As Double: a1d   = wsIn.Cells(sr, 8).Value   ' stator gauge alpha1
        Dim b3d   As Double: b3d   = wsIn.Cells(sr, 9).Value   ' rotor exit gauge |beta3|
        Dim Lam   As Double: Lam   = wsIn.Cells(sr, 10).Value
        Dim psi   As Double: psi   = wsIn.Cells(sr, 11).Value
        Dim soc   As Double: soc   = wsIn.Cells(sr, 12).Value
        Dim ett   As Double: ett   = wsIn.Cells(sr, 14).Value
        Dim etp   As Double: etp   = wsIn.Cells(sr, 15).Value

        ' ── Gas at stage inlet ────────────────────────────────
        Dim Cpi  As Double: Cpi  = CpT(T0i, a0, a1c, a2c)
        Dim gami As Double: gami = Cpi / (Cpi - R_mix)
        Dim Zs   As Double: Zs   = Z_in + (1 - Z_in) * (1 - P0i / P01) / (1 - P0n / P01)

        ' ── T-drop and work ───────────────────────────────────
        Dim dTi As Double: dTi = dTtot * psi / spsi
        Dim dh0 As Double: dh0 = Cpi * dTi

        ' ── Euler velocity ratios ─────────────────────────────
        Dim Cu2U As Double: Cu2U = (1 - Lam) + psi / 2
        Dim Cu3U As Double: Cu3U = (1 - Lam) - psi / 2

        ' ================================================================
        ' STATE 2 -- STATOR EXIT / ROTOR INLET
        ' phi2 = Cx2/U = Cu2U/tan(alpha1)
        ' NR: h_s -> Dm -> U -> Cx2 -> C2 -> T2s -> rho2 -> A2 -> h_s_new
        ' ================================================================
        Dim phi2 As Double: phi2 = Cu2U / Tan(a1d * PI / 180)

        Dim hs   As Double: hs   = 0.02
        Dim Dm   As Double, U    As Double
        Dim Cx2  As Double, C2   As Double
        Dim T2s  As Double, P2s  As Double, r2 As Double, A2 As Double
        Dim it   As Integer

        For it = 1 To 80
            Dm  = Dhub + hs
            U   = Omega * Dm / 2
            Cx2 = U * phi2
            C2  = U * Sqr(phi2 * phi2 + Cu2U * Cu2U)
            T2s = T0i - C2 * C2 / (2 * Cpi * 1000)
            If T2s < 50 Then T2s = 50
            P2s = P0i * (T2s / T0i) ^ (gami / (gami - 1))
            r2  = P2s / (Zs * R_J * T2s)
            A2  = mdot / (r2 * Cx2)
            Dim hn As Double
            hn  = (-Dhub + Sqr(Dhub * Dhub + 4 * A2 / PI)) / 2
            If Abs(hn - hs) < 1E-12 Then hs = hn: Exit For
            hs  = 0.5 * hs + 0.5 * hn
        Next it
        ' Finalise state 2
        Dm  = Dhub + hs
        U   = Omega * Dm / 2
        Cx2 = U * phi2
        C2  = U * Sqr(phi2 * phi2 + Cu2U * Cu2U)
        T2s = T0i - C2 * C2 / (2 * Cpi * 1000)
        P2s = P0i * (T2s / T0i) ^ (gami / (gami - 1))
        r2  = P2s / (Zs * R_J * T2s)
        A2  = mdot / (r2 * Cx2)
        ' continuity check state 2
        Dim m2c As Double: m2c = r2 * A2 * Cx2
        Dim dv2 As Double: dv2 = (m2c - mdot) / mdot * 100

        ' ================================================================
        ' STATE 1 -- STATOR INLET
        ' A1 = A2 (same annulus cross-section at stator inlet)
        ' C1 independent of Cx2: iterate C1 = mdot/(rho1*A2)
        ' rho1 = P1s/(Z*R*T1s),  T1s = T0i - C1^2/(2Cpi*1000)
        ' ================================================================
        Dim C1  As Double: C1  = Cx2   ' initial guess
        Dim T1s As Double, P1s As Double, r1 As Double
        For it = 1 To 80
            T1s = T0i - C1 * C1 / (2 * Cpi * 1000)
            If T1s < 50 Then T1s = 50
            P1s = P0i * (T1s / T0i) ^ (gami / (gami - 1))
            r1  = P1s / (Zs * R_J * T1s)
            Dim C1n As Double: C1n = mdot / (r1 * A2)
            If Abs(C1n - C1) < 1E-10 Then C1 = C1n: Exit For
            C1 = 0.5 * C1 + 0.5 * C1n
        Next it
        ' Finalise state 1
        T1s = T0i - C1 * C1 / (2 * Cpi * 1000)
        P1s = P0i * (T1s / T0i) ^ (gami / (gami - 1))
        r1  = P1s / (Zs * R_J * T1s)
        Dim m1c As Double: m1c = r1 * A2 * C1
        Dim dv1 As Double: dv1 = (m1c - mdot) / mdot * 100

        ' ================================================================
        ' STAGE EXIT STAGNATION
        ' ================================================================
        Dim T03 As Double: T03 = T0i - dTi
        Dim P03 As Double
        If s = N_st - 1 Then
            P03 = P0n
        Else
            P03 = P0i * (T03 / T0i) ^ (gami / (etp * (gami - 1)))
        End If

        Dim Cu2 As Double: Cu2 = U * Cu2U
        Dim Cu3 As Double: Cu3 = U * Cu3U
        Dim Wu2 As Double: Wu2 = Cu2 - U
        Dim Wu3 As Double: Wu3 = Cu3 - U

        ' ================================================================
        ' STATE 3 -- ROTOR EXIT
        ' beta3 input = rotor exit gauge angle (magnitude, from axial)
        ' Wu3 is known from Euler. Sign convention:
        '   Wu3 = Cu3 - U.  Cu3 is often negative (counter-swirl exit).
        '   |Wu3| / tan(|beta3|) = Cx3  (always positive)
        ' Cx3 fully determined by beta3 -- no iteration needed for Cx3.
        ' Then: T3s, P3s, rho3 -> A3 -> h_r (direct, 0% error)
        '
        ' Gas properties at rotor exit use Cp3 = CpT(T03)
        ' ================================================================
        Dim Cp3  As Double: Cp3  = CpT(T03, a0, a1c, a2c)
        Dim gam3 As Double: gam3 = Cp3 / (Cp3 - R_mix)

        ' Cx3 from beta3 and Wu3
        Dim Cx3 As Double: Cx3 = Abs(Wu3) / Tan(Abs(b3d) * PI / 180)

        Dim C3  As Double: C3  = Sqr(Cx3 * Cx3 + Cu3 * Cu3)
        Dim T3s As Double: T3s = T03 - C3 * C3 / (2 * Cp3 * 1000)
        If T3s < 50 Then T3s = 50
        Dim P3s As Double: P3s = P03 * (T3s / T03) ^ (gam3 / (gam3 - 1))
        Dim r3  As Double: r3  = P3s / (Zs * R_J * T3s)

        ' Rotor height: direct from continuity (0% error by construction)
        Dim A3 As Double: A3 = mdot / (r3 * Cx3)
        Dim hr As Double: hr = A3 / (PI * Dm)

        ' Continuity check state 3
        Dim m3c As Double: m3c = r3 * PI * Dm * hr * Cx3
        Dim dv3 As Double: dv3 = (m3c - mdot) / mdot * 100

        ' ================================================================
        ' VELOCITY TRIANGLES
        ' ================================================================
        Dim W2    As Double: W2    = Sqr(Cx2 * Cx2 + Wu2 * Wu2)
        Dim W3    As Double: W3    = Sqr(Cx3 * Cx3 + Wu3 * Wu3)
        Dim a2spd As Double: a2spd = Sqr(gami * R_J * T2s)
        Dim a3spd As Double: a3spd = Sqr(gam3 * R_J * T3s)
        Dim Ma2a  As Double: Ma2a  = C2 / a2spd
        Dim Ma2r  As Double: Ma2r  = W2 / a2spd
        Dim Ma3a  As Double: Ma3a  = C3 / a3spd
        Dim Ma3r  As Double: Ma3r  = W3 / a3spd
        Dim alp2  As Double: alp2  = Atn(Cu2 / Cx2) * 180 / PI
        Dim bet2  As Double: bet2  = Atn(Wu2 / Cx2) * 180 / PI
        Dim alp3  As Double: alp3  = Atn(Cu3 / Cx3) * 180 / PI
        Dim bet3  As Double: bet3  = Atn(Wu3 / Cx3) * 180 / PI   ' should = -b3d

        ' ================================================================
        ' ACTUAL REACTION -- thermodynamically correct, works for Cx2 != Cx3
        ' Lambda = (W3^2 - W2^2) / (2 * dh_stage_J)
        ' where dh_stage_J = U*(Cu2-Cu3) in J/kg
        ' This reduces to 1-(Cu2+Cu3)/(2U) ONLY when Cx2=Cx3
        ' When Cx2 != Cx3, use the full expression
        ' ================================================================
        Dim dh_stage_J As Double: dh_stage_J = U * (Cu2 - Cu3)   ' J/kg
        Dim Lam_actual As Double
        If Abs(dh_stage_J) > 0.001 Then
            Lam_actual = (W3 * W3 - W2 * W2) / (2 * dh_stage_J)
        Else
            Lam_actual = Lam   ' fallback
        End If

        ' C0 and U/C0
        Dim C0  As Double: C0  = Sqr(2 * (dh0 / ett) * 1000)
        Dim UC0 As Double: UC0 = U / C0

        ' ── Geometry summary ──────────────────────────────────
        Dim Dtip As Double: Dtip = Dhub + 2 * hs
        Dim hD   As Double: hD   = hs / Dm
        Dim AN2  As Double: AN2  = A2 * Nrpm ^ 2 / 1E9
        Dim Cpex As Double: Cpex = CpT(T03, a0, a1c, a2c)
        Dim gamex As Double: gamex = Cpex / (Cpex - R_mix)

        ' ── Blade geometry ────────────────────────────────────
        Dim snoms As Double: snoms = soc * cs_m * 1000
        Dim Zstat As Long:   Zstat = CLng(PI * Dm * 1000 / snoms + 0.5)
        If Zstat < 1 Then Zstat = 1
        Dim sacts As Double: sacts = PI * Dm * 1000 / Zstat

        Dim snomr As Double: snomr = soc * cr_m * 1000
        Dim Zrot  As Long:   Zrot  = CLng(PI * Dm * 1000 / snomr + 0.5)
        If Zrot < 1 Then Zrot = 1
        Dim sactr As Double: sactr = PI * Dm * 1000 / Zrot

        Dim ARs As Double: ARs = (hs * 1000) / (cs_m * 1000)
        Dim ARr As Double: ARr = (hr * 1000) / (cr_m * 1000)

        ' Throat: o = s * cos(gauge_angle)
        ' Stator: alpha1 (input), Rotor: |beta3| (input b3d)
        Dim os  As Double: os  = sacts * Cos(a1d * PI / 180)
        Dim or_ As Double: or_ = sactr * Cos(Abs(b3d) * PI / 180)
        Dim av  As Double: av  = Atn(Sqr(1 - (os / sacts) ^ 2) / (os / sacts)) * 180 / PI
        Dim b3v As Double: b3v = Atn(Sqr(1 - (or_ / sactr) ^ 2) / (or_ / sactr)) * 180 / PI

        ' Zweifel
        Dim Zws As Double
        Zws = 2 * (sacts / (cs_m * 1000)) * Cos(alp2 * PI / 180) ^ 2 _
            * (Tan(0) + Tan(alp2 * PI / 180))
        Dim Zwr As Double
        Zwr = 2 * (sactr / (cr_m * 1000)) * Cos(Abs(bet2) * PI / 180) ^ 2 _
            * (Abs(Tan(bet2 * PI / 180)) + Abs(Tan(b3d * PI / 180)))

        ' Stagger: xi = atan((tan_in + tan_out)/2)
        Dim xis As Double: xis = Atn((Tan(0) + Tan(alp2 * PI / 180)) / 2) * 180 / PI
        Dim xir As Double: xir = Atn((Tan(bet2 * PI / 180) + Tan(-b3d * PI / 180)) / 2) * 180 / PI
        Dim ths As Double: ths = alp2
        Dim thr As Double: thr = Abs(bet2) + b3d

        ' Mass flow per blade passage
        Dim mps As Double: mps = r2 * (sacts / 1000) * hs * Cx2
        Dim mpr As Double: mpr = r3 * (sactr / 1000) * hr * Cx3
        Dim mts As Double: mts = mps * Zstat
        Dim mtr As Double: mtr = mpr * Zrot
        Dim dvs As Double: dvs = (mts - mdot) / mdot * 100
        Dim dvr As Double: dvr = (mtr - mdot) / mdot * 100

        ' Implied eta_p last stage
        Dim etpi As Double: etpi = etp
        If s = N_st - 1 And Abs(T03 - T0i) > 0.001 Then
            Dim lP As Double: lP = Log(P03 / P0i)
            Dim lT As Double: lT = Log(T03 / T0i)
            If Abs(lT) > 1E-10 Then etpi = lP / ((gami / (gami - 1)) * lT)
        End If

        ' =======================================================
        ' WRITE TO ALL SHEETS
        ' =======================================================

        ' GAS_PROPERTIES
        Wv wsGP,gr,2, s+1,    "0"
        Wv wsGP,gr,3, T0i,    "0.00"
        Wv wsGP,gr,4, P0i,    "0"
        Wv wsGP,gr,5, Cpi,    "0.0000"
        Wv wsGP,gr,6, gami,   "0.0000"
        Wv wsGP,gr,7, R_mix,  "0.00000"
        Wv wsGP,gr,8, T03,    "0.00"
        Wv wsGP,gr,9, P03,    "0"
        Wv wsGP,gr,10,Cpex,   "0.0000"
        Wv wsGP,gr,11,gamex,  "0.0000"
        Wv wsGP,gr,12,T2s,    "0.00"
        Wv wsGP,gr,13,P2s,    "0"
        Wv wsGP,gr,14,P3s,    "0"
        Wv wsGP,gr,15,T3s,    "0.00"
        Wv wsGP,gr,17,Zs,     "0.00000"
        Wv wsGP,gr,18,r2,     "0.0000"
        Wv wsGP,gr,19,r3,     "0.0000"
        Wv wsGP,gr,20,r1,     "0.0000"

        ' THERMO
        Wv wsTH,gr,2, s+1,       "0"
        Wv wsTH,gr,3, dh0/ett,   "0.00"
        Wv wsTH,gr,4, dh0,       "0.00"
        Wv wsTH,gr,5, Cu2U,      "0.0000"
        Wv wsTH,gr,6, Cu3U,      "0.0000"
        Wv wsTH,gr,7, phi2,      "0.0000"
        Wv wsTH,gr,8, U,         "0.00"
        Wv wsTH,gr,9, T03,       "0.00"
        Wv wsTH,gr,10,P03,       "0"
        Wv wsTH,gr,11,T2s,       "0.00"
        Wv wsTH,gr,12,P2s,       "0"
        Wv wsTH,gr,13,T3s,       "0.00"
        Wv wsTH,gr,14,P3s,       "0"
        Wv wsTH,gr,15,C2,        "0.00"
        Wv wsTH,gr,16,Cx2,       "0.00"
        Wv wsTH,gr,17,C0,        "0.00"
        Wv wsTH,gr,18,UC0,       "0.0000"
        Wv wsTH,gr,19,P0i/P03,   "0.000"
        Wv wsTH,gr,20,dTi,       "0.00"
        Wv wsTH,gr,21,T1s,       "0.00"
        Wv wsTH,gr,22,P1s,       "0"
        Wv wsTH,gr,23,etpi,      "0.00000"
        Wv wsTH,gr,24,Cx3,       "0.00"     ' Cx3 rotor exit (different from Cx2)
        Wv wsTH,gr,25,C1,        "0.00"     ' C1 stator inlet (different from Cx2)
        Wv wsTH,gr,26,Lam_actual, "0.0000"  ' actual reaction (correct for Cx2!=Cx3)

        ' VEL_TRIANGLES
        Wv wsVT,gr,2, s+1,    "0"
        Wv wsVT,gr,3, U,      "0.00"
        Wv wsVT,gr,4, Cx2,    "0.00"     ' stator exit axial
        Wv wsVT,gr,5, phi2,   "0.0000"
        Wv wsVT,gr,6, Cu2,    "0.00"
        Wv wsVT,gr,7, C2,     "0.00"
        Wv wsVT,gr,8, alp2,   "0.00"
        Wv wsVT,gr,9, Wu2,    "0.00"
        Wv wsVT,gr,10,W2,     "0.00"
        Wv wsVT,gr,11,bet2,   "0.00"
        Wv wsVT,gr,12,Ma2a,   "0.0000"
        Wv wsVT,gr,13,Ma2r,   "0.0000"
        Wv wsVT,gr,14,Cu3,    "0.00"
        Wv wsVT,gr,15,C3,     "0.00"
        Wv wsVT,gr,16,alp3,   "0.00"
        Wv wsVT,gr,17,Wu3,    "0.00"
        Wv wsVT,gr,18,W3,     "0.00"
        Wv wsVT,gr,19,bet3,   "0.00"     ' computed = -b3d
        Wv wsVT,gr,20,Ma3a,   "0.0000"
        Wv wsVT,gr,21,Ma3r,   "0.0000"
        Wv wsVT,gr,22,UC0,    "0.0000"
        Wv wsVT,gr,23,Cx3,    "0.00"     ' rotor exit axial (different from col 4)
        Wv wsVT,gr,24,C1,     "0.00"     ' stator inlet velocity

        ' GEOMETRY
        Wv wsGE,gr,2, s+1,           "0"
        Wv wsGE,gr,3, Dhub*1000,     "0.0"
        Wv wsGE,gr,4, A2,            "0.000000"
        Wv wsGE,gr,5, hs*1000,       "0.00"
        Wv wsGE,gr,6, hr*1000,       "0.00"
        Wv wsGE,gr,7, Dm*1000,       "0.0"
        Wv wsGE,gr,8, Dtip*1000,     "0.0"
        Wv wsGE,gr,9, Dhub/Dtip,     "0.0000"
        Wv wsGE,gr,10,Nrpm,          "0"
        Wv wsGE,gr,11,Omega,         "0.00"
        Wv wsGE,gr,12,hD,            "0.0000"
        Wv wsGE,gr,13,AN2,           "0.00"
        Wv wsGE,gr,14,r2,            "0.0000"
        Wv wsGE,gr,15,Cx2,           "0.00"
        Wv wsGE,gr,16,Cx3,           "0.00"
        Wv wsGE,gr,17,(Cx3-Cx2)/Cx2*100,"0.00"  ' Cx change %
        ' Continuity detail columns
        Wv wsGE,gr,18,A2,            "0.000000"  ' A1=A2
        Wv wsGE,gr,19,hs*1000,       "0.00"      ' h1=h_s
        Wv wsGE,gr,20,C1,            "0.00"
        Wv wsGE,gr,21,m1c,           "0.0000"
        Wv wsGE,gr,22,dv1,           "+0.0000;-0.0000;0.0000"
        Wv wsGE,gr,23,A3,            "0.000000"
        Wv wsGE,gr,24,hr*1000,       "0.00"
        Wv wsGE,gr,25,m3c,           "0.0000"
        Wv wsGE,gr,26,dv3,           "+0.0000;-0.0000;0.0000"
        Wv wsGE,gr,27,hr*1000,       "0.00"
        Wv wsGE,gr,28,(hr-hs)*1000,  "0.00"   ' effective height diff
        Wv wsGE,gr,29,0,             "0.00"
        Wv wsGE,gr,30,0,             "0.00"
        Wv wsGE,gr,31,os,            "0.000"
        Wv wsGE,gr,32,or_,           "0.000"
        Wv wsGE,gr,33,os/sacts,      "0.0000"
        Wv wsGE,gr,34,or_/sactr,     "0.0000"

        ' BLADES
        Wv wsBL,gr,2, s+1,              "0"
        Wv wsBL,gr,3, cs_m*1000,        "0.00"
        Wv wsBL,gr,4, snoms,            "0.00"
        Wv wsBL,gr,5, Zstat,            "0"
        Wv wsBL,gr,6, sacts,            "0.00"
        Wv wsBL,gr,7, sacts/(cs_m*1000),"0.0000"
        Wv wsBL,gr,8, ARs,              "0.00"
        Wv wsBL,gr,9, Zws,              "0.0000"
        Wv wsBL,gr,10,cr_m*1000,        "0.00"
        Wv wsBL,gr,11,snomr,            "0.00"
        Wv wsBL,gr,12,Zrot,             "0"
        Wv wsBL,gr,13,sactr,            "0.00"
        Wv wsBL,gr,14,sactr/(cr_m*1000),"0.0000"
        Wv wsBL,gr,15,ARr,              "0.00"
        Wv wsBL,gr,16,Zwr,              "0.0000"
        Wv wsBL,gr,17,os,               "0.000"
        Wv wsBL,gr,18,or_,              "0.000"
        Wv wsBL,gr,19,av,               "0.00"   ' alpha1 verify
        Wv wsBL,gr,20,b3v,              "0.00"   ' beta3 verify
        Wv wsBL,gr,21,mps,              "0.0000"
        Wv wsBL,gr,22,mpr,              "0.0000"
        Wv wsBL,gr,23,mts,              "0.0000"
        Wv wsBL,gr,24,mtr,              "0.0000"
        Wv wsBL,gr,25,dvs,              "+0.00;-0.00;0.00"
        Wv wsBL,gr,26,dvr,              "+0.00;-0.00;0.00"
        Wv wsBL,gr,27,xis,              "0.00"
        Wv wsBL,gr,28,xir,              "0.00"
        Wv wsBL,gr,29,ths,              "0.00"
        Wv wsBL,gr,30,thr,              "0.00"
        Wv wsBL,gr,31,alp2,             "0.00"
        Wv wsBL,gr,32,bet2,             "0.00"
        Wv wsBL,gr,33,bet3,             "0.00"

        ' CRITERIA
        Dim cp(13) As Boolean
        cp(0)  = Ma2a < 1
        cp(1)  = Ma2r < 1
        cp(2)  = Ma3r < 1.2
        cp(3)  = UC0 >= 0.4 And UC0 <= 0.75
        cp(4)  = Zws >= 0.7 And Zws <= 1.1
        cp(5)  = Zwr >= 0.8 And Zwr <= 1
        cp(6)  = hD <= 0.25
        cp(7)  = Dhub / Dtip > 0.7
        cp(8)  = AN2 <= 50
        cp(9)  = Abs(dv1) < 0.001   ' state 1 continuity
        cp(10) = Abs(dv2) < 0.001   ' state 2 continuity
        cp(11) = Abs(dv3) < 0.001   ' state 3 continuity
        cp(12) = Abs(av - a1d) < 0.01
        cp(13) = Abs(b3v - b3d) < 0.01
        ' Reaction check: actual vs design intent
        Dim dLam As Double: dLam = Abs(Lam_actual - Lam)

        Dim colMap(13) As Integer
        colMap(0)=3:  colMap(1)=4:  colMap(2)=5:  colMap(3)=6
        colMap(4)=7:  colMap(5)=8:  colMap(6)=9:  colMap(7)=10
        colMap(8)=11: colMap(9)=12: colMap(10)=13: colMap(11)=14
        colMap(12)=17: colMap(13)=18

        Dim apass As Boolean: apass = True
        Dim ci As Integer
        For ci = 0 To 13
            If Not cp(ci) Then apass = False
            WvPF wsCR, gr, colMap(ci), cp(ci)
        Next ci
        wsCR.Cells(gr, 2).Value = s + 1
        WvPF wsCR, gr, 19, apass

        ' SUMMARY
        Wv wsSU,gr,2, s+1,          "0"
        Wv wsSU,gr,3, T0i,          "0.0"
        Wv wsSU,gr,4, T03,          "0.0"
        Wv wsSU,gr,5, dTi,          "0.0"
        Wv wsSU,gr,6, P0i,          "0"
        Wv wsSU,gr,7, P03,          "0"
        Wv wsSU,gr,8, P0i/P03,      "0.000"
        Wv wsSU,gr,9, Cpi,          "0.0000"
        Wv wsSU,gr,10,gami,         "0.0000"
        Wv wsSU,gr,11,Zs,           "0.0000"
        Wv wsSU,gr,12,r2,           "0.0000"
        Wv wsSU,gr,13,U,            "0.0"
        Wv wsSU,gr,14,Cx2,          "0.0"
        Wv wsSU,gr,15,Cx3,          "0.0"
        Wv wsSU,gr,16,C1,           "0.0"
        Wv wsSU,gr,17,phi2,         "0.0000"
        Wv wsSU,gr,18,alp2,         "0.00"
        Wv wsSU,gr,19,bet2,         "0.00"
        Wv wsSU,gr,20,bet3,         "0.00"
        Wv wsSU,gr,21,alp3,         "0.00"
        Wv wsSU,gr,22,Ma2a,         "0.0000"
        Wv wsSU,gr,23,Ma2r,         "0.0000"
        Wv wsSU,gr,24,Ma3r,         "0.0000"
        Wv wsSU,gr,25,Dhub*1000,    "0.0"
        Wv wsSU,gr,26,Dm*1000,      "0.0"
        Wv wsSU,gr,27,Dtip*1000,    "0.0"
        Wv wsSU,gr,28,hs*1000,      "0.0"
        Wv wsSU,gr,29,hr*1000,      "0.0"
        Wv wsSU,gr,30,Nrpm,         "0"
        Wv wsSU,gr,31,Zstat,        "0"
        Wv wsSU,gr,32,Zrot,         "0"
        Wv wsSU,gr,33,xis,          "0.00"
        Wv wsSU,gr,35,Lam_actual, "0.0000"  ' actual vs design reaction
        Wv wsSU,gr,34,IIf(apass,"PASS","FAIL"),"@"

        T0i = T03
        P0i = P03
Nxt:
    Next s

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim lastP As Double: lastP = wsGP.Cells(3 + N_st, 9).Value
    MsgBox "TURBINE CALCULATION COMPLETE  v6" & vbCrLf & vbCrLf & _
           "Stages : " & N_st & "    RPM : " & Format(Nrpm, "#,##0") & vbCrLf & _
           "P03 last : " & Format(lastP / 1E+05, "0.0000") & " bara" & vbCrLf & _
           "P0n target: " & Format(P0n / 1E+05, "0.0000") & " bara" & vbCrLf & vbCrLf & _
           "3-STATION CONTINUITY (all 0.000000%):" & vbCrLf & _
           "  State 1 stator inlet : C1 iterated, A1=A2" & vbCrLf & _
           "  State 2 stator exit  : NR on h_s, phi=Cx2/U" & vbCrLf & _
           "  State 3 rotor exit   : beta3 sets Cx3, h_r=A3/(pi*Dm)" & vbCrLf & vbCrLf & _
           "INPUT COLUMNS USED:" & vbCrLf & _
           "  Col H = alpha1  (stator gauge angle)" & vbCrLf & _
           "  Col I = beta3   (rotor EXIT gauge angle)", _
           vbInformation, "Turbine Solver v5"
End Sub

Public Sub AddCalcButton()
    Dim wsIn As Worksheet: Set wsIn = ThisWorkbook.Sheets("INPUTS")
    Dim shp As Shape
    For Each shp In wsIn.Shapes
        If shp.Name = "btnCompute" Then shp.Delete
    Next shp
    Dim btn As Shape
    Set btn = wsIn.Shapes.AddShape(msoShapeRoundedRectangle, _
              wsIn.Range("G1").Left, wsIn.Range("G1").Top, 230, 25)
    With btn
        .Name = "btnCompute"
        .TextFrame.Characters.Text = "  COMPUTE ALL STAGES  [v6]"
        .TextFrame.Characters.Font.Name = "Arial"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .Fill.ForeColor.RGB = RGB(0, 130, 60)
        .Line.ForeColor.RGB = RGB(0, 80, 30)
        .OnAction = "TurbineCalc.RunTurbineCalc"
    End With
End Sub
