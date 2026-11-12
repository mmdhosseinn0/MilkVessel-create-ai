Attribute VB_Name = "MilkVessel"
'==========================================================================
' PRESSURIZED MILK STORAGE VESSEL - SolidWorks VBA macro (parametric)
' Vertical vessel, 2:1 ellipsoidal heads, 4 legs, manways, sanitary nozzles
'
' Design basis (edit constants below):
'   ASME BPVC VIII-1 (UG-27 shell, UG-32 2:1 ellipsoidal head, UG-16 min t,
'   UG-36/37 opening check note, UG-99 hydrotest), PED 2014/68/EU,
'   3-A / EHEDG hygienic design, DIN 11850 tubes, EN 10088 316L.
'   Internal pressure only; vacuum protected by vacuum breaker (N4).
'
' Units: inputs in mm / MPa / kg/m3. SolidWorks API works in metres.
' HOW TO RUN: Tools > Macro > New..., in VBA editor File > Import File this
' .bas (or paste), press F5, run "main".
'==========================================================================
Option Explicit

Private Const PI As Double = 3.14159265358979

'------------------------- DESIGN INPUTS ---------------------------------
Private Const C_VOL_L As Double = 10000#   ' nominal total volume [L]
Private Const C_RI As Double = 1000#       ' inside radius [mm]  (ID 2000)
Private Const C_PDES As Double = 0.3       ' design pressure [MPa g] at top
Private Const C_TDES As Double = 150#      ' design temperature [deg C] (SIP)
Private Const C_S As Double = 115#         ' allowable stress 316L @T [MPa] - verify ASME II-D
Private Const C_SY As Double = 170#        ' yield 316L [MPa]
Private Const C_E As Double = 0.85         ' weld joint efficiency (spot RT)
Private Const C_CA As Double = 0.5         ' corrosion allowance [mm]
Private Const C_RHO As Double = 1035#      ' milk density [kg/m3]
Private Const C_TFAB As Double = 5#        ' min fabrication thickness [mm]
Private Const C_THIN As Double = 0.9       ' head forming thinning factor
Private Const C_YBOT As Double = 500#      ' elevation of bottom outer pole [mm]
Private Const C_LDRAIN As Double = 250#    ' outlet nozzle length below head [mm]
Private Const C_FILL As Double = 0.9       ' max fill ratio
Private Const C_LEGR As Double = 700#      ' leg pitch radius [mm]
Private Const C_LEGOD As Double = 114.3    ' leg tube OD [mm]
Private Const C_LEGWT As Double = 4#       ' leg tube wall [mm]
Private Const C_PADD As Double = 220#      ' foot pad diameter [mm]
Private Const C_PADT As Double = 12#       ' foot pad thickness [mm]

'------------------------- COMPUTED --------------------------------------
Private gTs As Double, gTh As Double, gRo As Double
Private gHh As Double, gHout As Double, gLcyl As Double
Private gY1 As Double, gY2 As Double, gYtop As Double
Private gTrS As Double, gTrH As Double, gPhyd As Double, gPbot As Double
Private gPallowBot As Double, gPallowTop As Double
Private gPtest As Double, gSigTest As Double, gVin As Double

Private Type NozzleT
    tag As String
    desc As String
    kind As Long      ' 1 top head, 2 shell +Z, 3 shell +X, 4 bottom pole
    px As Double      ' top: sketch X
    py As Double      ' top: sketch Y ; shell: height above lower tangent line
    od As Double
    wt As Double
    proj As Double    ' projection beyond surface [mm]
End Type

Private NZ(1 To 12) As NozzleT
Private swApp As SldWorks.SldWorks
Private swModel As SldWorks.ModelDoc2
Private swFM As SldWorks.FeatureManager
Private swSM As SldWorks.SketchManager

'==========================================================================
Sub main()
    On Error GoTo ErrH
    Dim tpl As String, i As Long
    Set swApp = Application.SldWorks

    CalcDesign
    DefineNozzles

    tpl = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
    If tpl = "" Then
        MsgBox "Default part template not set (Tools > Options > Default Templates).", vbExclamation
        Exit Sub
    End If
    Set swModel = swApp.NewDocument(tpl, 0, 0, 0)
    If swModel Is Nothing Then
        MsgBox "Could not create a new part.", vbCritical
        Exit Sub
    End If
    Set swFM = swModel.FeatureManager
    Set swSM = swModel.SketchManager
    swModel.Extension.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swUnitsLinear, 0, swLengthUnit_e.swMM

    SetMaterial
    BuildOuterSolid
    BuildLegs
    BuildNecks
    BuildCavity
    BuildBores
    swModel.ForceRebuild3 True
    SetProps
    swModel.ShowNamedView2 "*Isometric", 7
    swModel.ViewZoomtofit2

    Dim rp As String
    rp = WriteReport()
    MsgBox "Vessel generated." & vbCrLf & vbCrLf & _
           "Shell t = " & Format(gTs, "0.0") & " mm, Head t = " & Format(gTh, "0.0") & " mm" & vbCrLf & _
           "MAWP(top) = " & Format(gPallowTop, "0.000") & " MPa" & vbCrLf & _
           "Hydrotest = " & Format(gPtest, "0.000") & " MPa" & vbCrLf & _
           "Report: " & rp, vbInformation, "Milk pressure vessel"
    Exit Sub
ErrH:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'--------------------------- CALCULATIONS --------------------------------
Private Sub CalcDesign()
    Dim R As Double, Ps As Double, Ph As Double, Htot As Double
    gHh = C_RI / 2#                                   ' 2:1 head depth
    gLcyl = (C_VOL_L * 1000000# / (PI * C_RI * C_RI)) - (4# / 3#) * gHh
    gLcyl = Round(gLcyl / 10#, 0) * 10#               ' tangent-to-tangent
    Htot = gLcyl + 2# * gHh
    gPhyd = C_RHO * 9.81 * (Htot / 1000#) / 1000000#  ' MPa
    gPbot = C_PDES + gPhyd
    R = C_RI + C_CA
    gTrS = gPbot * R / (C_S * C_E - 0.6 * gPbot)                    ' UG-27
    gTrH = gPbot * 2# * R / (2# * C_S * C_E - 0.2 * gPbot)          ' UG-32(d)
    gTs = NextStd(MaxD(gTrS + C_CA, C_TFAB))
    gTh = NextStd(MaxD(gTrH / C_THIN + C_CA, gTs))
    gRo = C_RI + gTs
    gHout = gHh + gTh
    gY1 = C_YBOT + gHout          ' lower tangent line
    gY2 = gY1 + gLcyl             ' upper tangent line
    gYtop = gY2 + gHout
    Ps = C_S * C_E * (gTs - C_CA) / (R + 0.6 * (gTs - C_CA))
    Ph = 2# * C_S * C_E * (gTh - C_CA) / (2# * R + 0.2 * (gTh - C_CA))
    gPallowBot = MinD(Ps, Ph)
    gPallowTop = gPallowBot - gPhyd
    gPtest = 1.3 * C_PDES                                       ' UG-99(b), stress ratio 1
    gSigTest = gPtest * (C_RI + 0.6 * gTs) / gTs
    gVin = PI * C_RI * C_RI * gLcyl + 2# * (2# / 3#) * PI * C_RI * C_RI * gHh   ' mm3
End Sub

Private Function NextStd(t As Double) As Double
    Dim a As Variant, i As Long
    a = Array(3#, 4#, 5#, 6#, 8#, 10#, 12#, 15#, 20#, 25#)
    For i = LBound(a) To UBound(a)
        If a(i) >= t Then NextStd = a(i): Exit Function
    Next i
    NextStd = t
End Function
Private Function MaxD(a As Double, b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function
Private Function MinD(a As Double, b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function
Private Function MM(v As Double) As Double
    MM = v / 1000#
End Function

'--------------------------- NOZZLE TABLE --------------------------------
Private Sub SetNz(i As Long, tg As String, ds As String, k As Long, x As Double, y As Double, od As Double, wt As Double, pj As Double)
    NZ(i).tag = tg: NZ(i).desc = ds: NZ(i).kind = k
    NZ(i).px = x: NZ(i).py = y: NZ(i).od = od: NZ(i).wt = wt: NZ(i).proj = pj
End Sub

Private Sub DefineNozzles()
    SetNz 1, "N1", "Top manway DN500 (inward-opening cover)", 1, 540, 0, 508, 8, 300
    SetNz 2, "N2", "Agitator / mixer port DN250", 1, 0, 0, 273, 6, 250
    SetNz 3, "N3", "CIP spray ball DN65", 1, -400, 0, 70, 2, 200
    SetNz 4, "N4", "Safety valve + vacuum breaker DN80", 1, -750, 0, 85, 2, 200
    SetNz 5, "N5", "Pressure transmitter DN50", 1, 0, 500, 53, 1.5, 150
    SetNz 6, "N6", "Sterile air / steam inlet DN50", 1, 0, -500, 53, 1.5, 150
    SetNz 7, "N7", "Milk inlet DN65", 2, 0, 300, 70, 2, 150
    SetNz 8, "N8", "Side manway DN400", 2, 0, 1000, 406, 6, 150
    SetNz 9, "N9", "Temperature thermowell DN50", 2, 0, 1800, 53, 1.5, 120
    SetNz 10, "N10", "Level transmitter DN50", 3, 0, 500, 53, 1.5, 120
    SetNz 11, "N11", "CIP return / sampling DN50", 3, 0, 1500, 53, 1.5, 120
    SetNz 12, "N12", "Bottom outlet DN80 (self-draining)", 4, 0, 0, 85, 2, 0
End Sub

'--------------------------- SKETCH HELPERS ------------------------------
Private Function GetPlane(idx As Long) As SldWorks.Feature
    ' idx: 0 Front, 1 Top, 2 Right (default template order, language independent)
    Dim f As SldWorks.Feature, k As Long
    Set f = swModel.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then
            If k = idx Then Set GetPlane = f: Exit Function
            k = k + 1
        End If
        Set f = f.GetNextFeature
    Loop
End Function

Private Sub BeginSketch(planeIdx As Long)
    swModel.ClearSelection2 True
    GetPlane(planeIdx).Select2 False, 0
    swSM.InsertSketch True
    swSM.AddToDB = True
End Sub

Private Sub EndSketch()
    swSM.AddToDB = False
    swSM.InsertSketch True
End Sub

Private Sub SelLastFeat()
    Dim f As SldWorks.Feature
    swModel.ClearSelection2 True
    Set f = swModel.FeatureByPositionReverse(0)
    f.Select2 False, 0
End Sub

Private Sub AddEllipseSpline(cx As Double, cy As Double, a As Double, b As Double, ph1 As Double, ph2 As Double)
    Dim n As Long, i As Long, ph As Double
    Dim p() As Double
    n = 24
    ReDim p(0 To (n + 1) * 3 - 1)
    For i = 0 To n
        ph = ph1 + (ph2 - ph1) * i / n
        p(3 * i) = MM(cx + a * Cos(ph))
        p(3 * i + 1) = MM(cy + b * Sin(ph))
        p(3 * i + 2) = 0#
    Next i
    swSM.CreateSpline2 p, True
End Sub

Private Sub Ln(x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    swSM.CreateLine MM(x1), MM(y1), 0#, MM(x2), MM(y2), 0#
End Sub

Private Sub Circ(x As Double, y As Double, r As Double)
    swSM.CreateCircleByRadius MM(x), MM(y), 0#, MM(r)
End Sub

'--------------------------- FEATURE HELPERS -----------------------------
Private Sub Revolve360(isCut As Boolean, nm As String, ax As SldWorks.SketchSegment)
    Dim sd As SldWorks.SelectData, f As SldWorks.Feature
    SelLastFeat
    Set sd = swModel.SelectionManager.CreateSelectData
    sd.Mark = 4
    ax.Select4 True, sd
    Set f = swFM.FeatureRevolve2(True, True, False, isCut, False, False, 0, 0, 6.28318530717959, 0, False, False, 0, 0, 0, 0, 0, True, True, True)
    If f Is Nothing Then Err.Raise vbObjectError + 1, , "Revolve failed: " & nm
    f.Name = nm
End Sub

Private Sub ExtrudeBoss(nm As String, off As Double, depth As Double)
    Dim t0 As Long, f As SldWorks.Feature
    SelLastFeat
    t0 = swStartConditions_e.swStartSketchPlane
    If off <> 0 Then t0 = swStartConditions_e.swStartOffset
    Set f = swFM.FeatureExtrusion2(True, False, False, 0, 0, MM(depth), 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, t0, MM(off), False)
    If f Is Nothing Then Err.Raise vbObjectError + 2, , "Extrude failed: " & nm
    f.Name = nm
End Sub

Private Sub ExtrudeCut(nm As String, off As Double, depth As Double)
    Dim t0 As Long, f As SldWorks.Feature
    SelLastFeat
    t0 = swStartConditions_e.swStartSketchPlane
    If off <> 0 Then t0 = swStartConditions_e.swStartOffset
    Set f = swFM.FeatureCut3(True, False, False, 0, 0, MM(depth), 0, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, t0, MM(off), False, False)
    If f Is Nothing Then Err.Raise vbObjectError + 3, , "Cut failed: " & nm
    f.Name = nm
End Sub

'--------------------------- GEOMETRY ------------------------------------
Private Function YOuterBot(r As Double) As Double
    YOuterBot = C_YBOT + gHout - gHout * Sqr(1# - (r / gRo) ^ 2)
End Function
Private Function YOuterTop(r As Double) As Double
    YOuterTop = gY2 + gHout * Sqr(1# - (r / gRo) ^ 2)
End Function

Private Sub BuildOuterSolid()
    Dim ax As SldWorks.SketchSegment, z0 As Double
    BeginSketch 0
    AddEllipseSpline 0#, gY1, gRo, gHout, -PI / 2#, 0#
    Ln gRo, gY1, gRo, gY2
    AddEllipseSpline 0#, gY2, gRo, gHout, 0#, PI / 2#
    Ln 0#, gYtop, 0#, C_YBOT
    Set ax = swSM.CreateCenterLine(0#, MM(C_YBOT - 100#), 0#, 0#, MM(gYtop + 100#), 0#)
    EndSketch
    Revolve360 False, "Vessel_Outer_Solid", ax
End Sub

Private Sub BuildCavity()
    Dim ax As SldWorks.SketchSegment
    BeginSketch 0
    AddEllipseSpline 0#, gY1, C_RI, gHh, -PI / 2#, 0#
    Ln C_RI, gY1, C_RI, gY2
    AddEllipseSpline 0#, gY2, C_RI, gHh, 0#, PI / 2#
    Ln 0#, gY2 + gHh, 0#, gY1 - gHh
    Set ax = swSM.CreateCenterLine(0#, MM(gY1 - gHh - 100#), 0#, 0#, MM(gY2 + gHh + 100#), 0#)
    EndSketch
    Revolve360 True, "Vessel_Inner_Cavity_Cut", ax
End Sub

Private Sub BuildLegs()
    Dim k As Long, a As Double, x As Double, y As Double, yTop As Double
    yTop = YOuterBot(C_LEGR) + 25#
    BeginSketch 1
    For k = 0 To 3
        a = PI / 4# + k * PI / 2#
        x = C_LEGR * Cos(a): y = C_LEGR * Sin(a)
        Circ x, y, C_PADD / 2#
    Next k
    EndSketch
    ExtrudeBoss "Leg_Foot_Pads", 0#, C_PADT

    BeginSketch 1
    For k = 0 To 3
        a = PI / 4# + k * PI / 2#
        x = C_LEGR * Cos(a): y = C_LEGR * Sin(a)
        Circ x, y, C_LEGOD / 2#
        Circ x, y, C_LEGOD / 2# - C_LEGWT
    Next k
    EndSketch
    ExtrudeBoss "Support_Legs", 0#, yTop
End Sub

Private Sub BuildNecks()
    Dim i As Long, rc As Double, rn As Double, ro As Double, yT As Double, yN As Double
    For i = 1 To 12
        ro = NZ(i).od / 2#
        Select Case NZ(i).kind
        Case 1   ' top head, vertical
            rc = Sqr(NZ(i).px ^ 2 + NZ(i).py ^ 2)
            rn = MaxD(rc - ro, 0#)
            yT = YOuterTop(rn) + NZ(i).proj
            BeginSketch 1
            Circ NZ(i).px, NZ(i).py, ro
            EndSketch
            ExtrudeBoss NZ(i).tag & "_Neck", gY2, yT - gY2
        Case 2, 3  ' shell, horizontal
            yN = gY1 + NZ(i).py
            BeginSketch IIf(NZ(i).kind = 2, 0, 2)
            Circ 0#, yN, ro
            EndSketch
            ExtrudeBoss NZ(i).tag & "_Neck", 0#, gRo + NZ(i).proj
        Case 4   ' bottom outlet
            BeginSketch 1
            Circ NZ(i).px, NZ(i).py, ro
            EndSketch
            ExtrudeBoss NZ(i).tag & "_Neck", C_YBOT - C_LDRAIN, C_LDRAIN + gTh + 10#
        End Select
    Next i
End Sub

Private Sub BuildBores()
    Dim i As Long, rc As Double, rn As Double, ro As Double, ri As Double
    Dim yT As Double, yN As Double
    For i = 1 To 12
        ro = NZ(i).od / 2#
        ri = ro - NZ(i).wt
        Select Case NZ(i).kind
        Case 1
            rc = Sqr(NZ(i).px ^ 2 + NZ(i).py ^ 2)
            rn = MaxD(rc - ro, 0#)
            yT = YOuterTop(rn) + NZ(i).proj
            BeginSketch 1
            Circ NZ(i).px, NZ(i).py, ri
            EndSketch
            ExtrudeCut NZ(i).tag & "_Bore", gY2, yT - gY2
        Case 2, 3
            yN = gY1 + NZ(i).py
            BeginSketch IIf(NZ(i).kind = 2, 0, 2)
            Circ 0#, yN, ri
            EndSketch
            ExtrudeCut NZ(i).tag & "_Bore", 0#, gRo + NZ(i).proj
        Case 4
            BeginSketch 1
            Circ NZ(i).px, NZ(i).py, ri
            EndSketch
            ExtrudeCut NZ(i).tag & "_Bore", C_YBOT - C_LDRAIN, C_LDRAIN + gTh + 25#
        End Select
    Next i
End Sub

'--------------------------- PROPERTIES / REPORT -------------------------
Private Sub SetMaterial()
    On Error Resume Next
    Dim p As SldWorks.PartDoc
    Set p = swModel
    p.SetMaterialPropertyName2 "Default", "solidworks materials.sldmat", "AISI 316L Stainless Steel Sheet (SS)"
    On Error GoTo 0
End Sub

Private Sub SetProps()
    Dim cp As SldWorks.CustomPropertyManager
    Set cp = swModel.Extension.CustomPropertyManager("")
    cp.Add3 "Description", 30, "Pressurized milk storage vessel", 2
    cp.Add3 "Code", 30, "ASME VIII-1 / PED 2014/68/EU", 2
    cp.Add3 "Material", 30, "AISI 316L (EN 1.4404)", 2
    cp.Add3 "DesignPressure_MPa", 30, Format(C_PDES, "0.00"), 2
    cp.Add3 "DesignTemp_C", 30, Format(C_TDES, "0"), 2
    cp.Add3 "ShellThickness_mm", 30, Format(gTs, "0.0"), 2
    cp.Add3 "HeadThickness_mm", 30, Format(gTh, "0.0"), 2
    cp.Add3 "Volume_L", 30, Format(gVin / 1000000#, "0"), 2
    cp.Add3 "InternalFinish", 30, "Ra<=0.8 um, electropolished, welds ground flush", 2
End Sub

Private Function WriteReport() As String
    Dim path As String, fn As Integer, i As Long, s As String
    Dim mass As Double, mp As SldWorks.MassProperty
    On Error Resume Next
    Set mp = swModel.Extension.CreateMassProperty
    mass = mp.mass
    On Error GoTo 0
    path = swModel.GetPathName
    If path = "" Then
        path = Environ("USERPROFILE") & "\Documents\MilkVessel_DesignReport.txt"
    Else
        path = Left(path, InStrRev(path, ".")) & "DesignReport.txt"
    End If
    s = "PRESSURIZED MILK STORAGE VESSEL - DESIGN REPORT" & vbCrLf & String(60, "=") & vbCrLf
    s = s & "Codes: ASME VIII-1 (UG-16,27,32,36,37,99), PED 2014/68/EU, 3-A/EHEDG, DIN 11850" & vbCrLf
    s = s & "Material: AISI 316L; S=" & C_S & " MPa @ " & C_TDES & " C (verify ASME II-D); E=" & C_E & vbCrLf
    s = s & "Design P (top) = " & C_PDES & " MPa g + hydrostatic " & Format(gPhyd, "0.0000") & " MPa" & vbCrLf
    s = s & "Corrosion allowance = " & C_CA & " mm; min fabrication t = " & C_TFAB & " mm" & vbCrLf & vbCrLf
    s = s & "Required shell t (UG-27)  = " & Format(gTrS, "0.00") & " mm" & vbCrLf
    s = s & "Required head t (UG-32d)  = " & Format(gTrH, "0.00") & " mm (/" & C_THIN & " thinning)" & vbCrLf
    s = s & "Selected shell / head t   = " & Format(gTs, "0.0") & " / " & Format(gTh, "0.0") & " mm" & vbCrLf
    s = s & "Allowable P at bottom     = " & Format(gPallowBot, "0.000") & " MPa" & vbCrLf
    s = s & "Allowable P at top (MAWP) = " & Format(gPallowTop, "0.000") & " MPa" & vbCrLf
    If gPallowTop >= C_PDES Then s = s & "  -> OK (>= design pressure)" & vbCrLf Else s = s & "  -> NOT OK" & vbCrLf
    s = s & "Hydrotest (1.3 x P)       = " & Format(gPtest, "0.000") & " MPa; hoop stress " & Format(gSigTest, "0.0") & " MPa (limit 0.9Sy=" & Format(0.9 * C_SY, "0") & ")" & vbCrLf & vbCrLf
    s = s & "ID = " & 2 * C_RI & " mm; tangent length = " & gLcyl & " mm; head depth = " & gHh & " mm" & vbCrLf
    s = s & "Geometric volume = " & Format(gVin / 1000000#, "0") & " L; working (" & C_FILL * 100 & "%) = " & Format(gVin / 1000000# * C_FILL, "0") & " L" & vbCrLf
    s = s & "Overall height = " & Format(gYtop, "0") & " mm; empty mass (SolidWorks) = " & Format(mass, "0") & " kg" & vbCrLf & vbCrLf
    s = s & "NOZZLES (tube OD x wall, mm)" & vbCrLf
    For i = 1 To 12
        s = s & NZ(i).tag & "  " & NZ(i).desc & "  " & NZ(i).od & " x " & NZ(i).wt & vbCrLf
    Next i
    s = s & vbCrLf & "NOTES / CHECKS STILL REQUIRED:" & vbCrLf
    s = s & "1) Vessel is NOT designed for full vacuum: fit vacuum breaker (N4) or add stiffening rings (UG-28)." & vbCrLf
    s = s & "2) Opening reinforcement area replacement (UG-37) for N1/N8 and nozzle loads - check in PV Elite/Compress." & vbCrLf
    s = s & "3) Leg-to-head attachment: add wear/reinforcement pads, check local stress (WRC 107/537) and wind/seismic." & vbCrLf
    s = s & "4) Hygienic: all internal welds full-penetration, ground flush, Ra<=0.8 um; no dead legs (L/D<=2)." & vbCrLf
    s = s & "5) Hydrotest with chloride-free water (<25 ppm), drain and dry; pickle/passivate after fabrication." & vbCrLf
    s = s & "6) Final design must be verified and stamped by a qualified pressure-vessel engineer/notified body." & vbCrLf
    fn = FreeFile
    Open path For Output As #fn
    Print #fn, s
    Close #fn
    WriteReport = path
End Function
