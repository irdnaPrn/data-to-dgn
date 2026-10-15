Attribute VB_Name = "ElementideEksport"
Option Explicit

Public Sub Start()
    Dim table As Object, records As Collection, dialog As frmEksport, skipped As Long
    On Error GoTo Failed
    Set table = TooriistaAknad.CodeTable
    If table Is Nothing Then Exit Sub
    Set records = ScanModel(table, skipped)
    If records.Count = 0 Then
        MsgBox "Eksporditavaid celle ega sirgete loikudega jooni ei leitud.", vbInformation, "Eksport"
        Exit Sub
    End If
    Set dialog = New frmEksport
    dialog.Configure table, records, skipped
    dialog.Show vbModal
    Unload dialog
    Exit Sub
Failed:
    MsgBox "Eksport ebaonnestus: " & Err.Description, vbExclamation, "Eksport"
End Sub

Private Function ScanModel(ByVal table As Object, ByRef skipped As Long) As Collection
    Dim records As New Collection, scan As ElementEnumerator, criteria As New ElementScanCriteria
    Dim el As Element, record As clsElementKirje, kind As String, points() As Point3d
    Dim storedKind As String, code As String, numbers As Variant, matches As Collection
    ' Only top-level geometric features, never a cell's constituent graphics.
    criteria.ExcludeAllTypes
    criteria.IncludeType msdElementTypeCellHeader
    criteria.IncludeType msdElementTypeSharedCell
    criteria.IncludeType msdElementTypeLine
    criteria.IncludeType msdElementTypeLineString
    criteria.IncludeType msdElementTypeShape
    criteria.IncludeType msdElementTypeComplexString
    criteria.IncludeType msdElementTypeComplexShape
    criteria.IncludeType msdElementTypeEllipse
    Set scan = ActiveModelReference.Scan(criteria)
    Do While scan.MoveNext
        Set el = scan.Current
        If el.IsComponentElement Then GoTo NextElement
        ' Old importer marker circles have no stored point association: skip them.
        If el.IsEllipseElement And Not ElementAndmed.HasData(el) Then GoTo NextElement
        If Not GetPoints(el, points, kind) Then
            skipped = skipped + 1
            GoTo NextElement
        End If
        Set record = New clsElementKirje
        Set record.Source = el
        record.Kind = kind
        record.PointCount = UBound(points) - LBound(points) + 1
        If ElementAndmed.ReadData(el, storedKind, code, numbers) Then
            If storedKind <> kind Or UBound(numbers) - LBound(numbers) + 1 <> record.PointCount Then
                record.Issue = "Elemendi " & DLongToString(el.ID) & _
                    " tippude arv voi tuup on muutunud; salvestatud punktinumbreid ei saa kindlalt siduda."
            End If
            record.Code = code
            record.Numbers = numbers
        Else
            record.NeedsWrite = True
        End If
        If Len(record.Code) = 0 And kind <> "PUNKT" Then
            Set matches = MatchCodes(table, kind, el.Level.Name, Signature(el))
            If matches.Count = 1 Then
                record.Code = CStr(matches(1))
                record.NeedsWrite = True
            ElseIf matches.Count = 0 Then
                record.MatchInfo = "Kood puudub"
            Else
                record.MatchInfo = "Mitu vastet: " & JoinCollection(matches, ", ")
            End If
        End If
        records.Add record
NextElement:
    Loop
    Set ScanModel = records
End Function

Public Function Signature(ByVal el As Element) As String
    If el.IsCellElement Then
        Signature = el.AsCellElement.Name
    ElseIf el.IsSharedCellElement Then
        Signature = el.AsSharedCellElement.Name
    ElseIf Not el.IsEllipseElement Then
        If el.LineStyleSource = msdValueSourceLevel Then
            Signature = el.Level.ElementLineStyle.Name
        Else
            Signature = el.LineStyle.Name
        End If
    End If
End Function

Public Function MatchCodes(ByVal table As Object, ByVal kind As String, ByVal levelName As String, ByVal signatureName As String) As Collection
    Dim matches As New Collection, key As Variant, definition As Variant
    For Each key In table.Keys
        definition = table(key)
        If StrComp(CStr(definition(0)), kind, vbTextCompare) = 0 And _
            StrComp(CStr(definition(1)), levelName, vbTextCompare) = 0 And _
            StrComp(CStr(definition(2)), signatureName, vbTextCompare) = 0 Then matches.Add CStr(key)
    Next key
    Set MatchCodes = matches
End Function

Public Function GetPoints(ByVal el As Element, ByRef points() As Point3d, ByRef kind As String) As Boolean
    Dim children As ElementEnumerator, child As Element, part() As Point3d, partKind As String
    Dim i As Long, count As Long, storedKind As String, code As String, numbers As Variant
    kind = "": Erase points
    If el.IsCellElement Then
        ReDim points(0 To 0): points(0) = el.AsCellElement.Origin
        kind = "CELL"
    ElseIf el.IsSharedCellElement Then
        ReDim points(0 To 0): points(0) = el.AsSharedCellElement.Origin
        kind = "CELL"
    ElseIf el.IsLineElement Or el.IsShapeElement Then
        points = el.AsVertexList.GetVertices
        kind = "JOON"
    ElseIf el.IsComplexStringElement Or el.IsComplexShapeElement Then
        Set children = el.AsComplexElement.GetSubElements
        Do While children.MoveNext
            Set child = children.Current
            If Not GetPoints(child, part, partKind) Then Exit Function
            If partKind <> "JOON" Then Exit Function
            For i = LBound(part) To UBound(part)
                If count > 0 Then
                    If i = LBound(part) And Point3dEqual(points(count - 1), part(i)) Then GoTo NextPoint
                End If
                ReDim Preserve points(0 To count)
                points(count) = part(i): count = count + 1
NextPoint:
            Next i
        Loop
        If count < 2 Then Exit Function
        kind = "JOON"
    ElseIf el.IsEllipseElement Then
        If Not ElementAndmed.ReadData(el, storedKind, code, numbers) Then Exit Function
        If storedKind <> "PUNKT" Then Exit Function
        ReDim points(0 To 0): points(0) = el.AsEllipseElement.CenterPoint
        kind = "PUNKT"
    Else
        Exit Function
    End If
    GetPoints = True
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal separator As String) As String
    Dim i As Long, result As String
    For i = 1 To values.Count
        If i > 1 Then result = result & separator
        result = result & CStr(values(i))
    Next i
    JoinCollection = result
End Function

Public Sub AssignCode(ByVal records As Collection, ByVal groups As Object, ByVal code As String, ByVal table As Object)
    Dim record As clsElementKirje, definition As Variant
    If Not table.Exists(code) Then Err.Raise vbObjectError + 121, , "Kood puudub kooditabelis: " & code
    definition = table(code)
    ' Validate every selected group before changing any in-memory record.
    For Each record In records
        If groups.Exists(record.GroupKey) Then
            If Len(record.Issue) > 0 Then Err.Raise vbObjectError + 120, , record.Issue
            If record.Kind <> "PUNKT" And record.Kind <> CStr(definition(0)) Then _
                Err.Raise vbObjectError + 122, , "Valitud elemendi tuup ei sobi selle koodiga."
        End If
    Next record
    For Each record In records
        If groups.Exists(record.GroupKey) Then
            record.Code = code
            record.MatchInfo = ""
            record.NeedsWrite = True
        End If
    Next record
End Sub

Public Function BuildTable(ByVal records As Collection, ByVal groups As Object, ByVal fields As Variant, ByVal separator As String) As String
    Dim used As Object, record As clsElementKirje, points() As Point3d, kind As String
    Dim i As Long, j As Long, result As String, row As String, value As String, selected As New Collection
    Dim modelRef As Long
    modelRef = ActiveModelReference.MdlModelRefP
    For Each record In records
        If groups.Exists(record.GroupKey) Then
            If Len(record.Issue) > 0 Then Err.Raise vbObjectError + 120, , record.Issue
            If Len(record.Code) = 0 Then Err.Raise vbObjectError + 123, , "Maara esmalt valitud elementide puuduv voi mitmene kood."
            If record.Source.ModelReference.MdlModelRefP <> modelRef Then Err.Raise vbObjectError + 124, , "Aktiivne mudel on muutunud. Ava ekspordiaken uuesti."
            If record.NeedsWrite Or IsEmpty(record.Numbers) Then
                If record.Source.IsLocked Or record.Source.Level.IsLocked Then _
                    Err.Raise vbObjectError + 125, , "Punktiseose salvestamiseks vabasta elemendi voi leveli lukustus: " & record.Source.Level.Name
            End If
            selected.Add record
        End If
    Next record
    If selected.Count = 0 Then Err.Raise vbObjectError + 126, , "Vali vahemalt yks kood."
    Set used = ElementAndmed.UsedNumbers
    ' Also reserve staged numbers from a previous failed attempt in this window.
    For Each record In records
        If Not IsEmpty(record.Numbers) Then
            For i = LBound(record.Numbers) To UBound(record.Numbers)
                used(CStr(record.Numbers(i))) = True
            Next i
        End If
    Next record
    result = Join(fields, separator) & vbCrLf
    ' Stage every row before writing any new association into the drawing.
    For Each record In selected
        If Not GetPoints(record.Source, points, kind) Then Err.Raise vbObjectError + 127, , "Elemendi geomeetriat ei saa lugeda."
        If UBound(points) - LBound(points) + 1 <> record.PointCount Then _
            Err.Raise vbObjectError + 128, , "Elemendi tippude arv on muutunud. Ava ekspordiaken uuesti."
        If IsEmpty(record.Numbers) Then
            record.Numbers = ElementAndmed.NewNumbers(record.PointCount, used)
            record.NeedsWrite = True
        End If
        For i = LBound(points) To UBound(points)
            row = ""
            For j = LBound(fields) To UBound(fields)
                Select Case CStr(fields(j))
                    Case "pnr": value = CStr(record.Numbers(i - LBound(points) + LBound(record.Numbers)))
                    Case "x": value = Coordinate(points(i).Y)
                    Case "y": value = Coordinate(points(i).X)
                    Case "z": value = Coordinate(points(i).Z)
                    Case "kood": value = record.Code
                End Select
                If j > LBound(fields) Then row = row & separator
                row = row & QuoteField(value, separator)
            Next j
            result = result & row & vbCrLf
        Next i
    Next record
    For Each record In selected
        If record.NeedsWrite Then
            ElementAndmed.Store record.Source, record.Kind, record.Code, record.Numbers
            record.Source.Rewrite
            record.NeedsWrite = False
        End If
    Next record
    BuildTable = result
End Function

Public Function Coordinate(ByVal value As Double) As String
    Coordinate = Replace(Format$(value, "0.000"), ",", ".")
End Function

Public Function QuoteField(ByVal value As String, ByVal separator As String) As String
    If InStr(value, separator) > 0 Or InStr(value, Chr$(34)) > 0 Or InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then
        value = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
    End If
    QuoteField = value
End Function

Public Sub WriteFile(ByVal path As String, ByVal content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile path, 2
    stream.Close
End Sub
