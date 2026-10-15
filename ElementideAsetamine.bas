Attribute VB_Name = "ElementideAsetamine"
Option Explicit

Public Sub Start()
    Dim table As Object, dialog As frmAsetamine, code As String, definition As Variant
    Dim command As clsPlaceFeature, lib As String
    On Error GoTo Failed
    If Not ActiveModelReference.Is3D Then Err.Raise vbObjectError + 110, , "Ava 3D DGN."
    Set table = TooriistaAknad.CodeTable
    If table Is Nothing Then Exit Sub
    Set dialog = New frmAsetamine
    dialog.Configure table
    dialog.Show vbModal
    code = dialog.SelectedCode
    Unload dialog
    If Len(code) = 0 Then Exit Sub
    definition = table(code)
    If definition(0) = "CELL" Then
        lib = ValiImport.PickFile("Vali celliteek", "Celliteegid (*.cel)", "*.cel", "")
        If Len(lib) = 0 Then Exit Sub
        AttachCellLibrary lib
    End If
    Set command = New clsPlaceFeature
    command.Configure code, definition
    CommandState.StartPrimitive command
    Exit Sub
Failed:
    MsgBox "Asetamine ebaonnestus: " & Err.Description, vbExclamation, "Elementide asetamine"
End Sub

Public Function CellAt(ByVal definition As Variant, ByRef point As Point3d) As Element
    Set CellAt = CreateCellElement2(CStr(definition(2)), point, _
        Point3dFromXYZ(CDbl(definition(3)), CDbl(definition(3)), 1), True, Matrix3dIdentity)
End Function

Public Sub AddFeature(ByVal el As Element, ByVal kind As String, ByVal code As String, ByVal levelName As String)
    Dim used As Object, numbers As Variant, points() As Point3d, lev As Level
    Dim added As Boolean, heightAdded As Boolean, heightText As TextElement, message As String
    Dim center As Point3d
    On Error GoTo Failed
    Set used = ElementAndmed.UsedNumbers
    If kind = "CELL" Then
        numbers = ElementAndmed.NewNumbers(1, used)
    Else
        points = el.AsLineElement.GetVertices
        numbers = ElementAndmed.NewNumbers(UBound(points) - LBound(points) + 1, used)
    End If
    ElementAndmed.Store el, kind, code, numbers
    Set lev = ValiImport.EnsureLevel(levelName)
    If lev.IsLocked Then Err.Raise vbObjectError + 111, , "Level on lukustatud: " & levelName
    Set el.Level = lev
    ActiveModelReference.AddElement el
    added = True
    ValiImport.ApplyLevel el, lev
    If el.IsCellElement Then
        If UCase$(el.AsCellElement.Name) = "RING02" Then
            center = el.AsCellElement.Origin
            Set heightText = RingHeight(center)
            Set heightText.Level = lev
            ActiveModelReference.AddElement heightText
            heightAdded = True
            ValiImport.ApplyLevel heightText, lev
            KorgusePooramine.LinkTextToCell heightText, el
            heightText.Redraw msdDrawingModeNormal
        End If
    End If
    el.Redraw msdDrawingModeNormal
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    If heightAdded Then ActiveModelReference.RemoveElement heightText
    If added Then ActiveModelReference.RemoveElement el
    On Error GoTo 0
    Err.Raise vbObjectError + 112, "ElementideAsetamine.AddFeature", message
End Sub

Private Function RingHeight(ByRef center As Point3d) As TextElement
    Dim text As TextElement, style As TextStyle, fontItem As Font
    Set fontItem = ActiveDesignFile.Fonts.Item("ENGINEERING")
    Set text = CreateTextElement1(Nothing, Replace(Format$(center.Z, "0.00"), ",", "."), _
        Point3dFromXYZ(center.X + 0.3, center.Y - 0.5, center.Z), Matrix3dIdentity)
    Set style = text.TextStyle
    style.Height = 0.85: style.Width = 0.65
    style.Justification = msdTextJustificationLeftBottom
    Set style.Font = fontItem
    Set text.TextStyle = style
    Set RingHeight = text
End Function
