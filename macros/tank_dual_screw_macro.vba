Option Explicit

' SolidWorks VBA macro
' Creates:
' 1) Tank envelope using inch dimensions from the provided sketch
' 2) 0.25 in shell thickness
' 3) Two screw-conveyor centerlines and screw/motor placeholders
'
' NOTE:
' - Update TEMPLATE_PATH to a valid part template on your machine.
' - The geometry is parameterized so dimensions can be tuned quickly.

Dim swApp As SldWorks.SldWorks
Dim swPart As SldWorks.ModelDoc2
Dim swFeat As SldWorks.Feature
Dim ok As Boolean

Private Const TEMPLATE_PATH As String = "C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2024\\templates\\Part.prtdot"

' ========= Tank dimensions (inches) =========
Private Const TANK_WIDTH As Double = 26#
Private Const WALL_THICKNESS As Double = 0.25

Private Const H_TOTAL As Double = 230#
Private Const H_STEP As Double = 105#
Private Const STEP_RUN As Double = 65#
Private Const TOP_RUN As Double = 120#
Private Const CHAMFER_DROP As Double = 14#   ' 45° with CHAMFER_RUN = 14
Private Const CHAMFER_RUN As Double = 14#
Private Const RIGHT_DROP As Double = 44#
Private Const OUTLET_BACK As Double = 22#

' ========= Conveyor / motor dimensions =========
Private Const SCREW_OD As Double = 10#
Private Const SCREW_SHAFT_DIA As Double = 3#
Private Const MOTOR_DIA As Double = 12#
Private Const MOTOR_LEN As Double = 14#

Sub main()
    Dim errors As Long

    Set swApp = Application.SldWorks
    Set swPart = CreateNewPartDoc()

    If swPart Is Nothing Then
        MsgBox "Unable to create a new part document. Check TEMPLATE_PATH and verify your default part template is configured in Tools > Options > Default Templates.", vbCritical, "Macro Error"
        Exit Sub
    End If

    swApp.ActivateDoc3 swPart.GetTitle, True, 0, errors

    SetUnitsInches
    BuildTankProfileAndExtrude
    ShellTank

    ' Two screw conveyors following the sloped transfer and discharge leg.
    BuildStraightScrew "SCREW_1", 2#, 4#, 175#, 172#
    BuildStraightScrew "SCREW_2", 177#, 172#, 199#, 172#

    ' Motors at discharge/drive ends.
    AddMotor "MOTOR_LEFT", 2#, 4#, -1#, -1#
    AddMotor "MOTOR_RIGHT", 199#, 172#, 1#, 0#

    swPart.ViewZoomtofit2
    swPart.ForceRebuild3 False
End Sub

Private Function CreateNewPartDoc() As SldWorks.ModelDoc2
    Dim candidateTemplate As String

    ' 1) Try the explicit template path first.
    candidateTemplate = TEMPLATE_PATH
    If Len(candidateTemplate) > 0 And Len(Dir$(candidateTemplate)) > 0 Then
        Set CreateNewPartDoc = swApp.NewDocument(candidateTemplate, 0, 0#, 0#)
        If Not CreateNewPartDoc Is Nothing Then Exit Function
    End If

    ' 2) Try SolidWorks default part template through NewPart (most reliable fallback).
    Set CreateNewPartDoc = swApp.NewPart
    If Not CreateNewPartDoc Is Nothing Then Exit Function

    ' 3) Last attempt: try the user preference string value for default part template.
    ' Some environments expose this differently, so we keep this as a final fallback.
    candidateTemplate = swApp.GetUserPreferenceStringValue(13)
    If Len(candidateTemplate) > 0 And Len(Dir$(candidateTemplate)) > 0 Then
        Set CreateNewPartDoc = swApp.NewDocument(candidateTemplate, 0, 0#, 0#)
    End If
End Function

Private Sub SetUnitsInches()
    ' 0 = IPS in swUserPreferenceIntegerValue_e / swUnitSystem_e
    swPart.Extension.SetUserPreferenceInteger 72, 0, 0
End Sub

Private Sub BuildTankProfileAndExtrude()
    Dim xA As Double, yA As Double
    Dim xB As Double, yB As Double
    Dim xC As Double, yC As Double
    Dim xD As Double, yD As Double
    Dim xE As Double, yE As Double
    Dim xF As Double, yF As Double
    Dim xG As Double, yG As Double
    Dim xH As Double, yH As Double

    xA = 0#: yA = 0#
    xB = 0#: yB = H_STEP
    xC = STEP_RUN: yC = H_STEP
    xD = STEP_RUN: yD = H_TOTAL
    xE = STEP_RUN + TOP_RUN: yE = H_TOTAL
    xF = xE + CHAMFER_RUN: yF = H_TOTAL - CHAMFER_DROP
    xG = xF: yG = yF - RIGHT_DROP
    xH = xG - OUTLET_BACK: yH = yG

    swPart.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    swPart.SketchManager.InsertSketch True

    ' Closed profile following the screenshot geometry.
    swPart.SketchManager.CreateLine InToM(xA), InToM(yA), 0, InToM(xB), InToM(yB), 0
    swPart.SketchManager.CreateLine InToM(xB), InToM(yB), 0, InToM(xC), InToM(yC), 0
    swPart.SketchManager.CreateLine InToM(xC), InToM(yC), 0, InToM(xD), InToM(yD), 0
    swPart.SketchManager.CreateLine InToM(xD), InToM(yD), 0, InToM(xE), InToM(yE), 0
    swPart.SketchManager.CreateLine InToM(xE), InToM(yE), 0, InToM(xF), InToM(yF), 0
    swPart.SketchManager.CreateLine InToM(xF), InToM(yF), 0, InToM(xG), InToM(yG), 0
    swPart.SketchManager.CreateLine InToM(xG), InToM(yG), 0, InToM(xH), InToM(yH), 0
    swPart.SketchManager.CreateLine InToM(xH), InToM(yH), 0, InToM(xA), InToM(yA), 0

    swPart.SketchManager.InsertSketch True

    swPart.FeatureManager.FeatureExtrusion3 True, False, False, 0, 0, InToM(TANK_WIDTH), 0, _
        False, False, False, False, 0, 0, False, False, False, False, _
        True, True, True, 0, 0, False
End Sub

Private Sub ShellTank()
    ' Applies the requested 0.25 in wall thickness.
    ' SolidWorks API names vary by version/type library, so try both shell methods.
    Dim fm As Object

    swPart.ClearSelection2 True
    Set fm = swPart.FeatureManager
    Set swFeat = Nothing

    On Error Resume Next
    Set swFeat = CallByName(fm, "InsertFeatureShell", VbMethod, InToM(WALL_THICKNESS), False)
    If swFeat Is Nothing Then
        Set swFeat = CallByName(fm, "InsertShell", VbMethod, InToM(WALL_THICKNESS), False)
    End If
    On Error GoTo 0

    If swFeat Is Nothing Then
        MsgBox "Shell feature could not be created automatically. Please verify the profile is a closed solid and apply Shell = " & WALL_THICKNESS & " in manually.", vbExclamation, "Shell Warning"
    End If
End Sub

Private Sub BuildStraightScrew(ByVal prefix As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    ' Simplified screw representation:
    ' - Outer tube (OD)
    ' - Inner shaft
    ' For production use, replace with helical flight sweep if required.

    Dim dx As Double, dy As Double, lenIn As Double, ang As Double
    dx = x2 - x1
    dy = y2 - y1
    lenIn = Sqr(dx * dx + dy * dy)
    ang = Atn2(dy, dx)

    swPart.ClearSelection2 True
    swPart.Extension.SelectByID2 "Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    swPart.SketchManager.InsertSketch True
    swPart.SketchManager.CreateCircleByRadius InToM(x1), InToM(y1), 0, InToM(SCREW_OD / 2#)
    swPart.SketchManager.InsertSketch True

    swPart.FeatureManager.FeatureExtrusion3 True, False, False, 0, 0, InToM(lenIn), 0, _
        False, False, False, False, 0, 0, False, False, False, False, _
        True, True, True, 0, 0, False

    RotateLastFeatureZ ang

    swPart.ClearSelection2 True
    swPart.Extension.SelectByID2 "Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    swPart.SketchManager.InsertSketch True
    swPart.SketchManager.CreateCircleByRadius InToM(x1), InToM(y1), 0, InToM(SCREW_SHAFT_DIA / 2#)
    swPart.SketchManager.InsertSketch True

    swPart.FeatureManager.FeatureExtrusion3 True, False, False, 0, 0, InToM(lenIn), 0, _
        False, False, False, False, 0, 0, False, False, False, False, _
        True, True, True, 0, 0, False

    RotateLastFeatureZ ang
End Sub

Private Sub AddMotor(ByVal namePrefix As String, ByVal x As Double, ByVal y As Double, ByVal dirX As Double, ByVal dirY As Double)
    Dim ang As Double
    ang = Atn2(dirY, dirX)

    swPart.ClearSelection2 True
    swPart.Extension.SelectByID2 "Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    swPart.SketchManager.InsertSketch True
    swPart.SketchManager.CreateCircleByRadius InToM(x), InToM(y), 0, InToM(MOTOR_DIA / 2#)
    swPart.SketchManager.InsertSketch True

    swPart.FeatureManager.FeatureExtrusion3 True, False, False, 0, 0, InToM(MOTOR_LEN), 0, _
        False, False, False, False, 0, 0, False, False, False, False, _
        True, True, True, 0, 0, False

    RotateLastFeatureZ ang
End Sub

Private Sub RotateLastFeatureZ(ByVal angRad As Double)
    ' Rotates the most-recent feature around Z axis to align with conveyor direction.
    ' Keep this helper isolated; replacement with robust transform is straightforward.
    Dim swMathUtil As SldWorks.MathUtility
    Dim swMath As SldWorks.MathTransform
    Dim data(15) As Double

    Set swMathUtil = swApp.GetMathUtility

    data(0) = Cos(angRad): data(1) = -Sin(angRad): data(2) = 0
    data(3) = Sin(angRad): data(4) = Cos(angRad): data(5) = 0
    data(6) = 0: data(7) = 0: data(8) = 1
    data(9) = 0: data(10) = 0: data(11) = 0
    data(12) = 1: data(13) = 0: data(14) = 0: data(15) = 0

    Set swMath = swMathUtil.CreateTransform(data)
    swPart.Extension.SelectByID2 "Boss-Extrude*", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0
    swPart.Extension.TransformBy swMath.ArrayData
End Sub

Private Function InToM(ByVal inches As Double) As Double
    InToM = inches * 0.0254
End Function

Private Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    If x = 0# Then
        If y >= 0# Then
            Atn2 = 1.5707963267949
        Else
            Atn2 = -1.5707963267949
        End If
    ElseIf x > 0# Then
        Atn2 = Atn(y / x)
    Else
        If y >= 0# Then
            Atn2 = Atn(y / x) + 3.14159265358979
        Else
            Atn2 = Atn(y / x) - 3.14159265358979
        End If
    End If
End Function
