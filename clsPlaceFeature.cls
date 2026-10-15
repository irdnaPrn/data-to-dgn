VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPlaceFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPrimitiveCommandEvents

Private mCode As String, mDefinition As Variant, mPoints() As Point3d, mCount As Long
Private mStyle As LineStyle, mLevel As Level

Public Sub Configure(ByVal code As String, ByVal definition As Variant)
    mCode = code: mDefinition = definition
    Set mLevel = ValiImport.EnsureLevel(CStr(definition(1)))
    If mLevel.IsLocked Then Err.Raise vbObjectError + 113, , "Level on lukustatud: " & mLevel.Name
    If definition(0) = "JOON" Then Set mStyle = ValiImport.FindLineStyle(CStr(definition(2)))
End Sub

Private Sub IPrimitiveCommandEvents_Start()
    CommandState.EnableAccuSnap
    ShowCommand "Aseta " & mCode & " - " & CStr(mDefinition(4))
    Prompt
    CommandState.StartDynamics
End Sub

Private Sub Prompt()
    If mDefinition(0) = "CELL" Then
        ShowPrompt "Vali celli asukoht"
    ElseIf mCount = 0 Then
        ShowPrompt "Vali joone alguspunkt"
    Else
        ShowPrompt "Vali jargmine punkt; parem klops lopetab joone"
    End If
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim el As Element
    On Error GoTo Failed
    If mDefinition(0) = "CELL" Then
        Set el = ElementideAsetamine.CellAt(mDefinition, Point)
        ElementideAsetamine.AddFeature el, "CELL", mCode, CStr(mDefinition(1))
    Else
        If mCount > 0 Then
            If Point3dEqual(mPoints(mCount - 1), Point) Then Exit Sub
        End If
        ReDim Preserve mPoints(0 To mCount)
        mPoints(mCount) = Point
        mCount = mCount + 1
    End If
    Prompt
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Elementide asetamine"
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    Dim el As Element
    On Error GoTo Failed
    If mDefinition(0) = "CELL" Or mCount = 0 Then
        CommandState.StartDefaultCommand
        Exit Sub
    End If
    If mCount >= 2 Then
        Set el = CreateLineElement1(Nothing, mPoints)
        Set el.LineStyle = mStyle
        ElementideAsetamine.AddFeature el, "JOON", mCode, CStr(mDefinition(1))
    End If
    mCount = 0
    Erase mPoints
    Prompt
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Elementide asetamine"
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    Dim preview As Element, points() As Point3d, i As Long
    On Error GoTo IgnoreFrame
    If mDefinition(0) = "CELL" Then
        Set preview = ElementideAsetamine.CellAt(mDefinition, Point)
    Else
        If mCount = 0 Then Exit Sub
        If Point3dEqual(mPoints(mCount - 1), Point) Then Exit Sub
        ReDim points(0 To mCount)
        For i = 0 To mCount - 1
            points(i) = mPoints(i)
        Next i
        points(mCount) = Point
        Set preview = CreateLineElement1(Nothing, points)
        Set preview.LineStyle = mStyle
    End If
    Set preview.Level = mLevel
    preview.Color = -1
    preview.Redraw DrawMode
IgnoreFrame:
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
    CommandState.StopDynamics
    Erase mPoints
    mCount = 0
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub
