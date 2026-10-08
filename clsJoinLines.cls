VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsJoinLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPrimitiveCommandEvents

Private mFirst As Element

Private Sub IPrimitiveCommandEvents_Start()
    CommandState.SetLocateCursor
    ShowCommand "Joonte yhendamine"
    ShowPrompt "Vali esimene joon"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim located As Element
    On Error GoTo Failed
    Set located = CommandState.LocateElement(Point, View, True)
    If located Is Nothing Then
        ShowPrompt "Joont ei leitud. Klopsa joonel."
        Exit Sub
    End If
    If Not located.IsLineElement Then
        ShowPrompt "Vali tavaline joon voi linestring."
        Exit Sub
    End If
    If located.ModelReference.MdlModelRefP <> ActiveModelReference.MdlModelRefP Then
        ShowPrompt "Vali joon aktiivsest mudelist."
        Exit Sub
    End If
    If located.IsLocked Or located.Level.IsLocked Then
        ShowPrompt "Valitud joon voi selle level on lukustatud."
        Exit Sub
    End If
    If mFirst Is Nothing Then
        Set mFirst = located
        ActiveModelReference.SelectElement mFirst, True
        ShowPrompt "Vali teine joon"
        Exit Sub
    End If
    If DLongToString(mFirst.ID) = DLongToString(located.ID) Then
        ShowPrompt "Vali teine joon"
        Exit Sub
    End If
    ActiveModelReference.UnselectElement mFirst
    JoonteYhendamine.JoinPair mFirst, located
    Set mFirst = Nothing
    ShowPrompt "Vali esimene joon"
    Exit Sub
Failed:
    MsgBox "Joonte yhendamine ebaonnestus: " & Err.Description, vbExclamation, "Joonte yhendamine"
    ClearFirst
    ShowPrompt "Vali esimene joon"
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    If mFirst Is Nothing Then
        CommandState.StartDefaultCommand
    Else
        ClearFirst
        ShowPrompt "Vali esimene joon"
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
    ClearFirst
End Sub

Private Sub ClearFirst()
    On Error Resume Next
    If Not mFirst Is Nothing Then ActiveModelReference.UnselectElement mFirst
    Set mFirst = Nothing
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub
