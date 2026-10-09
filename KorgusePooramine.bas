Attribute VB_Name = "KorgusePooramine"
Option Explicit

Private Const LINK_APP As String = "DATA_TO_DGN_RING02_V1"
Private Const LEGACY_X As Double = 0.3
Private Const LEGACY_Y As Double = -0.5
Private Const MATCH_TOLERANCE As Double = 0.01

Public Sub Start()
    CommandState.StartPrimitive New clsRotateHeight
End Sub

Public Sub LinkTextToCell(ByVal heightText As Element, ByVal ringCell As Element, Optional ByVal writeNow As Boolean = True)
    Dim data(0 To 0) As XDatum
    data(0).Type = msdXDatumTypeString
    data(0).Value = DLongToString(ringCell.ID)
    heightText.SetXData LINK_APP, data
    If writeNow Then heightText.Rewrite
End Sub

Public Function IsHeightText(ByVal el As Element) As Boolean
    If Not el.IsTextElement Then Exit Function
    If UCase$(el.Level.Name) <> "KORGUS-EH2000" Then Exit Function
    If el.HasXData(LINK_APP) Then
        IsHeightText = True
    Else
        IsHeightText = Abs(el.AsTextElement.TextStyle.Height - 0.85) < 0.0001 And _
            Abs(el.AsTextElement.TextStyle.Width - 0.65) < 0.0001
    End If
End Function

Public Function FindRing(ByVal heightText As Element) As Element
    Dim data() As XDatum, linked As Element
    Dim scan As ElementEnumerator, criteria As New ElementScanCriteria
    Dim candidate As Element, found As Element
    Dim textOrigin As Point3d, expectedCenter As Point3d, center As Point3d
    Dim rotation As Matrix3d, angle As Double, heightValue As String
    If heightText.HasXData(LINK_APP) Then
        data = heightText.GetXData(LINK_APP)
        If UBound(data) <> LBound(data) Then Err.Raise vbObjectError + 50, , "Korgusteksti RING02 seos on vigane."
        On Error Resume Next
        Set linked = ActiveModelReference.GetElementByID(DLongFromString(CStr(data(LBound(data)).Value)))
        On Error GoTo 0
        If linked Is Nothing Then Err.Raise vbObjectError + 51, , "Seotud RING02 celli ei leitud."
        If Not IsRing(linked) Then Err.Raise vbObjectError + 52, , "Seotud element ei ole RING02."
        Set FindRing = linked
        Exit Function
    End If

    ' Legacy text: recover the center from its original offset and XY rotation.
    textOrigin = heightText.AsTextElement.Origin
    rotation = heightText.AsTextElement.Rotation
    If Not Matrix3dIsXYRotation(rotation, angle) Then Err.Raise vbObjectError + 53, , "Tekst peab asuma XY-tasandis."
    expectedCenter = Point3dFromXYZ(textOrigin.X - (LEGACY_X * Cos(angle) - LEGACY_Y * Sin(angle)), _
        textOrigin.Y - (LEGACY_X * Sin(angle) + LEGACY_Y * Cos(angle)), textOrigin.Z)
    criteria.ExcludeAllTypes
    criteria.IncludeType msdElementTypeCellHeader
    Set scan = ActiveModelReference.Scan(criteria)
    Do While scan.MoveNext
        Set candidate = scan.Current
        If IsRing(candidate) Then
            center = candidate.AsCellElement.Origin
            heightValue = Replace(Format$(center.Z, "0.00"), ",", ".")
            If heightValue = Trim$(heightText.AsTextElement.Text) And _
                Abs(center.X - expectedCenter.X) <= MATCH_TOLERANCE And _
                Abs(center.Y - expectedCenter.Y) <= MATCH_TOLERANCE And _
                Abs(center.Z - expectedCenter.Z) <= MATCH_TOLERANCE Then
                If Not found Is Nothing Then Err.Raise vbObjectError + 54, , "Tekstile vastab mitu RING02 celli. Yhest vastet ei leitud."
                Set found = candidate
            End If
        End If
    Loop
    If found Is Nothing Then Err.Raise vbObjectError + 55, , "Tekstile vastavat RING02 celli ei leitud."
    Set FindRing = found
End Function

Private Function IsRing(ByVal el As Element) As Boolean
    If Not el.IsCellElement Then Exit Function
    IsRing = UCase$(el.AsCellElement.Name) = "RING02"
End Function

Public Function RayAngle(ByRef center As Point3d, ByRef point As Point3d) As Double
    Dim dx As Double, dy As Double, pi As Double
    dx = point.X - center.X
    dy = point.Y - center.Y
    pi = 4 * Atn(1)
    If dx = 0 And dy = 0 Then Err.Raise vbObjectError + 56, , "Vali suund celli keskpunktist eemal."
    If dx = 0 Then
        If dy > 0 Then RayAngle = pi / 2 Else RayAngle = -pi / 2
    Else
        RayAngle = Atn(dy / dx)
        If dx < 0 Then
            If dy >= 0 Then RayAngle = RayAngle + pi Else RayAngle = RayAngle - pi
        End If
    End If
End Function

Public Function RotatedCopy(ByVal original As Element, ByRef center As Point3d, ByVal angle As Double) As Element
    Dim result As Element, rotation As Matrix3d, transform As Transform3d
    rotation = Matrix3dFromAxisAndRotationAngle(2, angle)
    transform = Transform3dFromMatrix3dAndFixedPoint3d(rotation, center)
    Set result = original.Clone
    result.Transform transform
    Set RotatedCopy = result
End Function
