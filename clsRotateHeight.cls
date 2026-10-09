VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsRotateHeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPrimitiveCommandEvents

Private mText As Element, mRing As Element
Private mOriginal As Element
Private mCenter As Point3d
Private mStartAngle As Double

Private Sub IPrimitiveCommandEvents_Start()
    CommandState.SetLocateCursor
    ShowCommand "Korgusarvu pooramine"
    ShowPrompt "Vali korgusarv"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim located As Element, rotated As Element, angle As Double
    On Error GoTo Failed
    If mText Is Nothing Then
        Set located = CommandState.LocateElement(Point, View, True)
        If located Is Nothing Then Exit Sub
        If located.ModelReference.MdlModelRefP <> ActiveModelReference.MdlModelRefP Then Exit Sub
        If Not KorgusePooramine.IsHeightText(located) Then
            ShowPrompt "Vali suur korgusarv levelilt KORGUS-EH2000"
            Exit Sub
        End If
        If located.IsLocked Or located.Level.IsLocked Then
            ShowPrompt "Korgustekst voi selle level on lukustatud"
            Exit Sub
        End If
        Set mRing = KorgusePooramine.FindRing(located)
        mCenter = mRing.AsCellElement.Origin
        mStartAngle = KorgusePooramine.RayAngle(mCenter, Point)
        Set mOriginal = located.Clone
        Set mText = located
        ActiveModelReference.SelectElement mText, True
        CommandState.StartDynamics
        ShowPrompt "Vali poordenurk"
        Exit Sub
    End If

    If Point.X = mCenter.X And Point.Y = mCenter.Y Then Exit Sub
    angle = KorgusePooramine.RayAngle(mCenter, Point) - mStartAngle
    Set rotated = KorgusePooramine.RotatedCopy(mOriginal, mCenter, angle)
    KorgusePooramine.LinkTextToCell rotated, mRing, False
    CommandState.StopDynamics
    ActiveModelReference.UnselectElement mText
    ActiveModelReference.ReplaceElement mText, rotated
    rotated.Redraw msdDrawingModeNormal
    ClearText
    ShowPrompt "Vali korgusarv"
    Exit Sub
Failed:
    MsgBox "Korgusarvu pooramine ebaonnestus: " & Err.Description, vbExclamation, "Korgusarvu pooramine"
    ClearText
    ShowPrompt "Vali korgusarv"
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    Dim preview As Element, angle As Double
    On Error GoTo IgnoreFrame
    If mText Is Nothing Then Exit Sub
    If Point.X = mCenter.X And Point.Y = mCenter.Y Then Exit Sub
    angle = KorgusePooramine.RayAngle(mCenter, Point) - mStartAngle
    Set preview = KorgusePooramine.RotatedCopy(mOriginal, mCenter, angle)
    preview.Redraw DrawMode
IgnoreFrame:
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    If mText Is Nothing Then
        CommandState.StartDefaultCommand
    Else
        ClearText
        ShowPrompt "Vali korgusarv"
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
    ClearText
End Sub

Private Sub ClearText()
    On Error Resume Next
    CommandState.StopDynamics
    If Not mText Is Nothing Then ActiveModelReference.UnselectElement mText
    Set mText = Nothing
    Set mRing = Nothing
    Set mOriginal = Nothing
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub
