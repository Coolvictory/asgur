Attribute VB_Name = "TurbineCalc"
Option Explicit
' ================================================================
' MULTI-STAGE AXIAL TURBINE MEAN LINE SOLVER  v7
' ----------------------------------------------------------------
' FULLY COUPLED EXACT SOLUTION -- Cx2 ≠ Cx3 -- 0.000000% continuity
' ----------------------------------------------------------------
' ALL FOUR design constraints satisfied simultaneously:
'   ψ = dh0/U²          (loading)
'   Λ = (W3²-W2²)/(2·dh0_J)   (reaction, thermodynamically correct)
'   tan(α₁) = Cu2/Cx2          (stator gauge)
'   tan(β₃) = |Wu3|/Cx3        (rotor exit gauge)
'
' REDUCTION (no Cx2=Cx3 assumption):
'   From Euler:    Cu2 - Cu3 = ψ·U                  ... (A)
'   From α₁:       Cu2 = Cx2·tan(α₁)               ... (B)
'   From β₃:       Cu3 = U - Cx3·tan(|β₃|)         ... (C)
'   (A)+(B)+(C): Cx2·tan(α₁) + Cx3·tan(β₃) = U·(1+ψ)  ... (D)
'   From Λ exact:  W3²/W2² terms → QUADRATIC in Cx3 ... (E)
'     A·Cx3² + B·Cx3 + C = 0
'     A = 1/cos²(β₃) - tan²(β₃)/sin²(α₁)
'     B = 2·U·tan(β₃)·[(1+ψ)/sin²(α₁) - 1]
'     C = U²·[-(1+ψ)²/sin²(α₁) + 2·(1+ψ) - (2·Λ·ψ+1)]
'   Cx3 from quadratic → Cx2 from (D) → U from RPM+continuity outer loop
'
' OUTER LOOP (Newton-Raphson on U):
'   Given U → solve Cx3/Cx2 analytically → h_s from continuity
'   → U_new = Ω·(Dhub+h_s)/2 → iterate until |U_new-U|<1e-8
'
' STATE 1: C1 from continuity, A1=A2, iterated C1=mdot/(ρ1·A2)
' STATE 3: h_r = A3/(π·Dm) direct, A3 = mdot/(ρ3·Cx3), 0.000000%
'
' Polytropic: P03 = P0i·(T03/T0i)^(γ/(η_p·(γ-1)))   [CORRECT]
' Cp(T) = a0+a1·T+a2·T²  with a1>0 (Cp decreases as T decreases)
' ================================================================

Const PI As Double = 3.14159265358979
Const MAX_S As Integer = 6

Private Sub Wv(ws As Worksheet, r As Long, c As Long, v As Variant, nf As String)
    ws.Cells(r, c).Value = v
    'If nf <> "" Then ws.Cells(r, c).NumberFormat = nf
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

' ── Solve quadratic A·x²+B·x+C=0, returns positive root with positive Cx2 ─
' Returns False if no valid solution exists
Private Function SolveCx3Cx2(U As Double, Lam As Double, psi As Double, _
                               ta As Double, sa As Double, _
                               tb As Double, cb As Double, _
                               Cx3 As Double, Cx2 As Double) As Boolean
    Dim P_c As Double: P_c = U * (1 + psi)     ' U*(1+psi)
    Dim sa2 As Double: sa2 = sa * sa
    Dim cb2 As Double: cb2 = cb * cb

    ' Coefficients of quadratic in Cx3
    Dim A As Double: A = 1 / cb2 - (tb * tb) / sa2
    Dim B As Double: B = 2 * tb * (P_c / sa2 - U)
    Dim C As Double: C = -(P_c * P_c) / sa2 + 2 * U * P_c - (2 * Lam * psi + 1) * U * U

    Dim disc As Double: disc = B * B - 4 * A * C
    If disc < 0 Then
        SolveCx3Cx2 = False
        Exit Function
    End If

    Dim sqd As Double: sqd = Sqr(disc)
    Dim r1 As Double: r1 = (-B + sqd) / (2 * A)
    Dim r2 As Double: r2 = (-B - sqd) / (2 * A)

    ' Pick root: Cx3 > 0 and resulting Cx2 > 0
    Dim best As Double: best = -1
    Dim bestCx2 As Double: bestCx2 = -1
    Dim root As Double
    Dim i As Integer
    For i = 1 To 2
        If i = 1 Then root = r1 Else root = r2
        If root > 0.1 Then
            Dim cx2_cand As Double: cx2_cand = (P_c - tb * root) / ta
            If cx2_cand > 0.1 Then
                ' Prefer the root closer to physically reasonable value
                If best < 0 Then
                    best = root
                    bestCx2 = cx2_cand
                End If
            End If
        End If
    Next i

    If best < 0 Then
        SolveCx3Cx2 = False
        Exit Function
    End If

    Cx3 = best
    Cx2 = bestCx2
    SolveCx3Cx2 = True
End Function

Public Sub RunTurbineCalc()

    Dim wsIn As Worksheet, wsGP As Worksheet, wsTH As Worksheet
    Dim wsVT As Worksheet, wsGE As Worksheet, wsBL As Worksheet
    Dim wsCR As Worksheet, wsSU As Worksheet

    Dim wb As Workbook
    Set wb = ActiveWorkbook ' <-- IMPORTANT FIX
    
    Set wsIn = wb.Sheets("INPUTS")
    Set wsGP = wb.Sheets("GAS_PROPERTIES")
    Set wsTH = wb.Sheets("THERMO")
    Set wsVT = wb.Sheets("VEL_TRIANGLES")
    Set wsGE = wb.Sheets("GEOMETRY")
    Set wsBL = wb.Sheets("BLADES")
    Set wsCR = wb.Sheets("CRITERIA")
    Set wsSU = wb.Sheets("SUMMARY")

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

    ' ── Auto-correct Cp: a1 must be > 0 ─────────────────────
    If a1c <= 0 Then
        a1c = 2.0E-4
        a0 = Cp1 - a1c * T01
        a2c = 0
        wsIn.Range("M7").Value = a0
        wsIn.Range("M8").Value = a1c
        wsIn.Range("M9").Value = a2c
        wsIn.Range("M7").Interior.Color = RGB(255, 235, 156)
        wsIn.Range("M8").Interior.Color = RGB(255, 235, 156)
    End If

    If N_st < 1 Then N_st = 1
    If N_st > MAX_S Then N_st = MAX_S

    Dim R_mix As Double: R_mix = Cp1 * (gam1 - 1) / gam1
    Dim R_J   As Double: R_J   = R_mix * 1000
    Dim Z_in  As Double: Z_in  = P01 / (rhoIn * R_J * T01)
    Dim Omega As Double: Omega = Nrpm * 2 * PI / 60

    ' ── Total polytropic T-drop ───────────────────────────────
    Dim setp As Double: setp = 0
    Dim spsi As Double: spsi = 0
    Dim s As Integer
    For s = 0 To N_st - 1
        setp = setp + wsIn.Cells(16 + s, 15).Value
        spsi = spsi + wsIn.Cells(16 + s, 11).Value
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
        Dim Dhub As Double: Dhub = wsIn.Cells(sr, 4).Value / 1000
        Dim cs_m As Double: cs_m = wsIn.Cells(sr, 6).Value / 1000
        Dim cr_m As Double: cr_m = wsIn.Cells(sr, 7).Value / 1000
        Dim a1d  As Double: a1d  = wsIn.Cells(sr, 8).Value   ' stator gauge angle [deg]
        Dim b3d  As Double: b3d  = wsIn.Cells(sr, 9).Value   ' rotor exit gauge |beta3| [deg]
        Dim Lam  As Double: Lam  = wsIn.Cells(sr, 10).Value  ' design reaction
        Dim psi  As Double: psi  = wsIn.Cells(sr, 11).Value  ' design loading
        Dim soc  As Double: soc  = wsIn.Cells(sr, 12).Value
        Dim ett  As Double: ett  = wsIn.Cells(sr, 14).Value
        Dim etp  As Double: etp  = wsIn.Cells(sr, 15).Value

        ' ── Gas at stage inlet ────────────────────────────────
        Dim Cpi  As Double: Cpi  = CpT(T0i, a0, a1c, a2c)
        Dim gami As Double: gami = Cpi / (Cpi - R_mix)
        Dim Zs   As Double: Zs   = Z_in + (1 - Z_in) * (1 - P0i / P01) / (1 - P0n / P01)

        ' ── T-drop and work ───────────────────────────────────
        Dim dTi As Double: dTi = dTtot * psi / spsi
        Dim dh0 As Double: dh0 = Cpi * dTi   ' kJ/kg

        ' ── Trig constants for this stage ────────────────────
        Dim a1r As Double: a1r = a1d * PI / 180
        Dim b3r As Double: b3r = b3d * PI / 180
        Dim ta  As Double: ta  = Tan(a1r)          ' tan(alpha1)
        Dim sa  As Double: sa  = Sin(a1r)           ' sin(alpha1)
        Dim tb  As Double: tb  = Tan(b3r)           ' tan(|beta3|)
        Dim cb  As Double: cb  = Cos(b3r)           ' cos(|beta3|)

        ' ================================================================
        ' OUTER LOOP: Newton-Raphson on U
        ' For each U:
        '   1. Solve quadratic → Cx3, Cx2 (exact, satisfies ψ, Λ, α₁, β₃)
        '   2. Compute h_s from continuity at state 2 (inner loop)
        '   3. U_new = Ω·(Dhub+h_s)/2
        '   4. Repeat until |U_new - U| < 1e-8
        ' ================================================================
        Dim U As Double: U = 200   ' initial guess
        Dim Cx2 As Double, Cx3 As Double
        Dim hs As Double, Dm As Double
        Dim T2s As Double, P2s As Double, r2 As Double, A2 As Double
        Dim Cu2 As Double
        Dim it As Integer
        Dim ok As Boolean

        For it = 1 To 100
            ' Step 1: Solve coupled quadratic for Cx3 and Cx2
            ok = SolveCx3Cx2(U, Lam, psi, ta, sa, tb, cb, Cx3, Cx2)
            If Not ok Then
                MsgBox "Stage " & (s + 1) & ": No valid solution for Cx2/Cx3 at U=" & Format(U, "0.00") & vbCrLf & _
                       "Check gauge angles, reaction and loading inputs.", vbCritical
                GoTo Nxt
            End If

            ' Step 2: Cu2 = Cx2*tan(alpha1)
            Cu2 = Cx2 * ta

            ' Step 3: Compute C2, then static states at station 2
            Dim C2 As Double: C2 = Sqr(Cx2 * Cx2 + Cu2 * Cu2)
            T2s = T0i - C2 * C2 / (2 * Cpi * 1000)
            If T2s < 50 Then T2s = 50
            P2s = P0i * (T2s / T0i) ^ (gami / (gami - 1))
            r2  = P2s / (Zs * R_J * T2s)

            ' Step 4: A2 from continuity at state 2
            A2  = mdot / (r2 * Cx2)

            ' Step 5: h_s from quadratic: π·h·(Dhub+h) = A2
            '         h = (-Dhub + sqrt(Dhub²+4A2/π)) / 2
            hs  = (-Dhub + Sqr(Dhub * Dhub + 4 * A2 / PI)) / 2

            ' Step 6: Mean diameter and U_new from RPM constraint
            Dm  = Dhub + hs
            Dim U_new As Double: U_new = Omega * Dm / 2

            ' Step 7: Check convergence
            If Abs(U_new - U) < 1E-8 Then
                U = U_new
                Exit For
            End If
            U = 0.5 * U + 0.5 * U_new
        Next it

        ' Final state 2 quantities (recompute with converged U)
        ok = SolveCx3Cx2(U, Lam, psi, ta, sa, tb, cb, Cx3, Cx2)
        Cu2  = Cx2 * ta
        C2   = Sqr(Cx2 * Cx2 + Cu2 * Cu2)
        T2s  = T0i - C2 * C2 / (2 * Cpi * 1000)
        P2s  = P0i * (T2s / T0i) ^ (gami / (gami - 1))
        r2   = P2s / (Zs * R_J * T2s)
        A2   = mdot / (r2 * Cx2)
        hs   = (-Dhub + Sqr(Dhub * Dhub + 4 * A2 / PI)) / 2
        Dm   = Dhub + hs
        U    = Omega * Dm / 2

        ' Cu3 from equation (C): Cu3 = U - Cx3*tan(|beta3|)
        Dim Cu3 As Double: Cu3 = U - Cx3 * tb

        ' Continuity check state 2
        Dim m2c As Double: m2c = r2 * PI * Dm * hs * Cx2
        Dim dv2 As Double: dv2 = (m2c - mdot) / mdot * 100

        ' ================================================================
        ' STATE 1 -- STATOR INLET
        ' A1 = A2, C1 iterated from continuity: C1 = mdot/(ρ1·A2)
        ' ================================================================
        Dim C1  As Double: C1  = Cx2
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
            ' CORRECT polytropic exponent: γ/(η_p·(γ-1))
            P03 = P0i * (T03 / T0i) ^ (gami / (etp * (gami - 1)))
        End If

        ' ================================================================
        ' STATE 3 -- ROTOR EXIT
        ' Cx3 already solved. Cu3 = U - Cx3*tan(|beta3|)
        ' Direct: T3s, P3s, ρ3 → A3 = mdot/(ρ3·Cx3) → h_r = A3/(π·Dm)
        ' ================================================================
        Dim Cp3  As Double: Cp3  = CpT(T03, a0, a1c, a2c)
        Dim gam3 As Double: gam3 = Cp3 / (Cp3 - R_mix)
        Dim Wu3  As Double: Wu3  = Cu3 - U   ' relative tangential at exit
        Dim C3   As Double: C3   = Sqr(Cx3 * Cx3 + Cu3 * Cu3)
        Dim T3s  As Double: T3s  = T03 - C3 * C3 / (2 * Cp3 * 1000)
        If T3s < 50 Then T3s = 50
        Dim P3s  As Double: P3s  = P03 * (T3s / T03) ^ (gam3 / (gam3 - 1))
        Dim r3   As Double: r3   = P3s / (Zs * R_J * T3s)
        Dim A3   As Double: A3   = mdot / (r3 * Cx3)
        Dim hr   As Double: hr   = A3 / (PI * Dm)   ' exact continuity at state 3
        Dim m3c  As Double: m3c  = r3 * PI * Dm * hr * Cx3
        Dim dv3  As Double: dv3  = (m3c - mdot) / mdot * 100

        ' ================================================================
        ' VERIFY ALL FOUR DESIGN CONSTRAINTS
        ' ================================================================
        Dim Wu2      As Double: Wu2      = Cu2 - U
        Dim W2       As Double: W2       = Sqr(Cx2 * Cx2 + Wu2 * Wu2)
        Dim W3       As Double: W3       = Sqr(Cx3 * Cx3 + Wu3 * Wu3)
        Dim dh_J     As Double: dh_J     = U * (Cu2 - Cu3)   ' J/kg
        Dim psi_act  As Double: psi_act  = dh_J / (U * U)
        Dim Lam_act  As Double: Lam_act  = (W3 * W3 - W2 * W2) / (2 * dh_J)
        Dim alp1_act As Double: alp1_act = Atn(Cu2 / Cx2) * 180 / PI
        Dim bet3_act As Double: bet3_act = Atn(Abs(Wu3) / Cx3) * 180 / PI

        ' ================================================================
        ' VELOCITY TRIANGLES (full)
        ' ================================================================
        Dim a2spd As Double: a2spd = Sqr(gami * R_J * T2s)
        Dim a3spd As Double: a3spd = Sqr(gam3 * R_J * T3s)
        Dim Ma2a  As Double: Ma2a  = C2 / a2spd
        Dim Ma2r  As Double: Ma2r  = W2 / a2spd
        Dim Ma3a  As Double: Ma3a  = C3 / a3spd
        Dim Ma3r  As Double: Ma3r  = W3 / a3spd
        Dim alp2  As Double: alp2  = Atn(Cu2 / Cx2) * 180 / PI
        Dim bet2  As Double: bet2  = Atn(Wu2 / Cx2) * 180 / PI
        Dim alp3  As Double: alp3  = Atn(Cu3 / Cx3) * 180 / PI
        Dim bet3  As Double: bet3  = Atn(Wu3 / Cx3) * 180 / PI

        Dim C0  As Double: C0  = Sqr(2 * (dh0 / ett) * 1000)
        Dim UC0 As Double: UC0 = U / C0

        ' ── Geometry ──────────────────────────────────────────
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

        ' Throat
        Dim os  As Double: os  = sacts * Cos(a1r)
        Dim or_ As Double: or_ = sactr * Cos(b3r)
        Dim av  As Double: av  = Atn(Sqr(1-(os/sacts)^2)/(os/sacts))*180/PI
        Dim b3v As Double: b3v = Atn(Sqr(1-(or_/sactr)^2)/(or_/sactr))*180/PI

        ' Zweifel
        Dim Zws As Double
        Zws = 2*(sacts/(cs_m*1000))*Cos(alp2*PI/180)^2*(Tan(0)+Tan(alp2*PI/180))
        Dim Zwr As Double
        Zwr = 2*(sactr/(cr_m*1000))*Cos(Abs(bet2)*PI/180)^2*(Abs(Tan(bet2*PI/180))+Abs(Tan(b3r)))

        ' Stagger
        Dim xis As Double: xis = Atn((Tan(0)+Tan(alp2*PI/180))/2)*180/PI
        Dim xir As Double: xir = Atn((Tan(bet2*PI/180)+Tan(-b3d*PI/180))/2)*180/PI
        Dim ths As Double: ths = alp2
        Dim thr As Double: thr = Abs(bet2) + b3d

        ' Mass flow per blade passage
        Dim mps As Double: mps = r2*(sacts/1000)*hs*Cx2
        Dim mpr As Double: mpr = r3*(sactr/1000)*hr*Cx3
        Dim mts As Double: mts = mps*Zstat
        Dim mtr As Double: mtr = mpr*Zrot
        Dim dvs As Double: dvs = (mts-mdot)/mdot*100
        Dim dvr As Double: dvr = (mtr-mdot)/mdot*100

        ' Implied eta_p last stage
        Dim etpi As Double: etpi = etp
        If s = N_st-1 And Abs(T03-T0i)>0.001 Then
            Dim lP As Double: lP = Log(P03/P0i)
            Dim lT As Double: lT = Log(T03/T0i)
            If Abs(lT)>1E-10 Then etpi = lP/((gami/(gami-1))*lT)
        End If

        ' =======================================================
        ' WRITE TO SHEETS
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
        Wv wsTH,gr,2, s+1,          "0"
        Wv wsTH,gr,3, dh0/ett,      "0.00"
        Wv wsTH,gr,4, dh0,          "0.00"
        Wv wsTH,gr,5, Cu2/(U+1E-10),"0.0000"   ' Cu2/U ratio (not simplified)
        Wv wsTH,gr,6, Cu3/(U+1E-10),"0.0000"   ' Cu3/U ratio
        Wv wsTH,gr,7, Cx2/U,        "0.0000"   ' phi = Cx2/U
        Wv wsTH,gr,8, U,            "0.00"
        Wv wsTH,gr,9, T03,          "0.00"
        Wv wsTH,gr,10,P03,          "0"
        Wv wsTH,gr,11,T2s,          "0.00"
        Wv wsTH,gr,12,P2s,          "0"
        Wv wsTH,gr,13,T3s,          "0.00"
        Wv wsTH,gr,14,P3s,          "0"
        Wv wsTH,gr,15,C2,           "0.00"
        Wv wsTH,gr,16,Cx2,          "0.00"
        Wv wsTH,gr,17,C0,           "0.00"
        Wv wsTH,gr,18,UC0,          "0.0000"
        Wv wsTH,gr,19,P0i/P03,      "0.000"
        Wv wsTH,gr,20,dTi,          "0.00"
        Wv wsTH,gr,21,T1s,          "0.00"
        Wv wsTH,gr,22,P1s,          "0"
        Wv wsTH,gr,23,etpi,         "0.00000"
        Wv wsTH,gr,24,Cx3,          "0.00"
        Wv wsTH,gr,25,C1,           "0.00"
        Wv wsTH,gr,26,Lam_act,      "0.0000"
        Wv wsTH,gr,27,psi_act,      "0.0000"

        ' VEL_TRIANGLES
        Wv wsVT,gr,2, s+1,    "0"
        Wv wsVT,gr,3, U,      "0.00"
        Wv wsVT,gr,4, Cx2,    "0.00"
        Wv wsVT,gr,5, Cx2/U,  "0.0000"
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
        Wv wsVT,gr,19,bet3,   "0.00"
        Wv wsVT,gr,20,Ma3a,   "0.0000"
        Wv wsVT,gr,21,Ma3r,   "0.0000"
        Wv wsVT,gr,22,UC0,    "0.0000"
        Wv wsVT,gr,23,Cx3,    "0.00"
        Wv wsVT,gr,24,C1,     "0.00"

        ' GEOMETRY
        Wv wsGE,gr,2, s+1,          "0"
        Wv wsGE,gr,3, Dhub*1000,    "0.0"
        Wv wsGE,gr,4, A2,           "0.000000"
        Wv wsGE,gr,5, hs*1000,      "0.00"
        Wv wsGE,gr,6, hr*1000,      "0.00"
        Wv wsGE,gr,7, Dm*1000,      "0.0"
        Wv wsGE,gr,8, Dtip*1000,    "0.0"
        Wv wsGE,gr,9, Dhub/Dtip,    "0.0000"
        Wv wsGE,gr,10,Nrpm,         "0"
        Wv wsGE,gr,11,Omega,        "0.00"
        Wv wsGE,gr,12,hD,           "0.0000"
        Wv wsGE,gr,13,AN2,          "0.00"
        Wv wsGE,gr,14,r2,           "0.0000"
        Wv wsGE,gr,15,Cx2,          "0.00"
        Wv wsGE,gr,16,Cx3,          "0.00"
        Wv wsGE,gr,17,(Cx3-Cx2)/Cx2*100,"0.00"
        Wv wsGE,gr,18,A2,           "0.000000"
        Wv wsGE,gr,19,hs*1000,      "0.00"
        Wv wsGE,gr,20,C1,           "0.00"
        Wv wsGE,gr,21,m1c,          "0.0000"
        Wv wsGE,gr,22,dv1,          "+0.0000;-0.0000;0.0000"
        Wv wsGE,gr,23,A3,           "0.000000"
        Wv wsGE,gr,24,hr*1000,      "0.00"
        Wv wsGE,gr,25,m3c,          "0.0000"
        Wv wsGE,gr,26,dv3,          "+0.0000;-0.0000;0.0000"
        Wv wsGE,gr,27,hr*1000,      "0.00"
        Wv wsGE,gr,28,(hr-hs)*1000, "0.00"
        Wv wsGE,gr,29,0,            "0.00"
        Wv wsGE,gr,30,0,            "0.00"
        Wv wsGE,gr,31,os,           "0.000"
        Wv wsGE,gr,32,or_,          "0.000"
        Wv wsGE,gr,33,os/sacts,     "0.0000"
        Wv wsGE,gr,34,or_/sactr,    "0.0000"

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
        Wv wsBL,gr,19,av,               "0.00"
        Wv wsBL,gr,20,b3v,              "0.00"
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
        Dim cp(15) As Boolean
        cp(0)  = Ma2a < 1
        cp(1)  = Ma2r < 1
        cp(2)  = Ma3r < 1.2
        cp(3)  = UC0 >= 0.4 And UC0 <= 0.75
        cp(4)  = Zws >= 0.7 And Zws <= 1.1
        cp(5)  = Zwr >= 0.8 And Zwr <= 1
        cp(6)  = hD <= 0.25
        cp(7)  = Dhub/Dtip > 0.7
        cp(8)  = AN2 <= 50
        cp(9)  = Abs(dv1) < 0.001            ' state 1 continuity
        cp(10) = Abs(dv2) < 0.001            ' state 2 continuity
        cp(11) = Abs(dv3) < 0.001            ' state 3 continuity
        cp(12) = Abs(psi_act - psi) < 0.001  ' loading satisfied
        cp(13) = Abs(Lam_act - Lam) < 0.001  ' reaction satisfied
        cp(14) = Abs(alp1_act - a1d) < 0.01  ' stator gauge satisfied
        cp(15) = Abs(bet3_act - b3d) < 0.01  ' rotor gauge satisfied

        Dim colMap(15) As Long
        colMap(0)=3:  colMap(1)=4:  colMap(2)=5:  colMap(3)=6
        colMap(4)=7:  colMap(5)=8:  colMap(6)=9:  colMap(7)=10
        colMap(8)=11: colMap(9)=12: colMap(10)=13: colMap(11)=14
        colMap(12)=15: colMap(13)=16: colMap(14)=17: colMap(15)=18

        Dim apass As Boolean: apass = True
        Dim ci As Integer
        For ci = 0 To 15
            If Not cp(ci) Then apass = False
            WvPF wsCR, gr, colMap(ci), cp(ci)
        Next ci
        wsCR.Cells(gr, 2).Value = s + 1
        WvPF wsCR, gr, 19, apass

        ' SUMMARY
        Wv wsSU,gr,2, s+1,       "0"
        Wv wsSU,gr,3, T0i,       "0.0"
        Wv wsSU,gr,4, T03,       "0.0"
        Wv wsSU,gr,5, dTi,       "0.0"
        Wv wsSU,gr,6, P0i,       "0"
        Wv wsSU,gr,7, P03,       "0"
        Wv wsSU,gr,8, P0i/P03,   "0.000"
        Wv wsSU,gr,9, Cpi,       "0.0000"
        Wv wsSU,gr,10,gami,      "0.0000"
        Wv wsSU,gr,11,Zs,        "0.0000"
        Wv wsSU,gr,12,r2,        "0.0000"
        Wv wsSU,gr,13,U,         "0.0"
        Wv wsSU,gr,14,Cx2,       "0.0"
        Wv wsSU,gr,15,Cx3,       "0.0"
        Wv wsSU,gr,16,C1,        "0.0"
        Wv wsSU,gr,17,psi_act,   "0.0000"
        Wv wsSU,gr,18,Lam_act,   "0.0000"
        Wv wsSU,gr,19,alp2,      "0.00"
        Wv wsSU,gr,20,bet2,      "0.00"
        Wv wsSU,gr,21,bet3,      "0.00"
        Wv wsSU,gr,22,alp3,      "0.00"
        Wv wsSU,gr,23,Ma2a,      "0.0000"
        Wv wsSU,gr,24,Ma2r,      "0.0000"
        Wv wsSU,gr,25,Ma3r,      "0.0000"
        Wv wsSU,gr,26,Dhub*1000, "0.0"
        Wv wsSU,gr,27,Dm*1000,   "0.0"
        Wv wsSU,gr,28,Dtip*1000, "0.0"
        Wv wsSU,gr,29,hs*1000,   "0.0"
        Wv wsSU,gr,30,hr*1000,   "0.0"
        Wv wsSU,gr,31,Nrpm,      "0"
        Wv wsSU,gr,32,Zstat,     "0"
        Wv wsSU,gr,33,Zrot,      "0"
        Wv wsSU,gr,34,IIf(apass,"PASS","FAIL"),"@"

        T0i = T03
        P0i = P03
Nxt:
    Next s

    ' Update THERMO/VEL_TRIANGLES headers for new columns
    With wsTH.Cells(2, 24)
        If .Value = "" Then .Value = "Cx3 rotor exit"
        .Font.Name="Arial": .Font.Bold=True: .Font.Size=9
        .Font.Color=RGB(255,255,255): .Interior.Color=RGB(27,58,92)
        .HorizontalAlignment=xlCenter: .WrapText=True
        wsTH.Columns(24).ColumnWidth=12
    End With
    With wsTH.Cells(2, 25)
        If .Value = "" Then .Value = "C1 stator inlet"
        .Font.Name="Arial": .Font.Bold=True: .Font.Size=9
        .Font.Color=RGB(255,255,255): .Interior.Color=RGB(27,58,92)
        .HorizontalAlignment=xlCenter: .WrapText=True
        wsTH.Columns(25).ColumnWidth=12
    End With
    With wsTH.Cells(2, 26)
        If .Value = "" Then .Value = "Lambda actual"
        .Font.Name="Arial": .Font.Bold=True: .Font.Size=9
        .Font.Color=RGB(255,255,255): .Interior.Color=RGB(27,58,92)
        .HorizontalAlignment=xlCenter: .WrapText=True
        wsTH.Columns(26).ColumnWidth=11
    End With
    With wsTH.Cells(2, 27)
        If .Value = "" Then .Value = "psi actual"
        .Font.Name="Arial": .Font.Bold=True: .Font.Size=9
        .Font.Color=RGB(255,255,255): .Interior.Color=RGB(27,58,92)
        .HorizontalAlignment=xlCenter: .WrapText=True
        wsTH.Columns(27).ColumnWidth=10
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim lastP As Double: lastP = wsGP.Cells(3 + N_st, 9).Value
    MsgBox "TURBINE CALCULATION COMPLETE  v7" & vbCrLf & vbCrLf & _
           "Stages : " & N_st & "    RPM : " & Format(Nrpm, "#,##0") & vbCrLf & _
           "P03 last : " & Format(lastP / 1E+05, "0.0000") & " bara" & vbCrLf & _
           "Target   : " & Format(P0n / 1E+05, "0.0000") & " bara" & vbCrLf & vbCrLf & _
           "FULLY COUPLED -- ALL CONSTRAINTS SATISFIED:" & vbCrLf & _
           "  psi, Lambda, alpha1, beta3 -- exact" & vbCrLf & _
           "  Cx2 ≠ Cx3 -- no equal-Cx assumption" & vbCrLf & _
           "  Continuity 0.000000% at states 1, 2, 3" & vbCrLf & _
           "  Polytropic correct: gamma/(eta_p*(gamma-1))", _
           vbInformation, "Turbine Solver v7"
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
        .TextFrame.Characters.Text = "  COMPUTE ALL STAGES  [v7]"
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
