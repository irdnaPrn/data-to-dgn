Attribute VB_Name = "ValiImport"
Option Explicit

' Windows Unicode file dialog, for 32-bit PowerDraft V8i / VBA6.
Private Type OpenFileInfo
    size As Long
    owner As Long
    instance As Long
    filter As Long
    customFilter As Long
    maxCustomFilter As Long
    filterIndex As Long
    file As Long
    maxFile As Long
    fileTitle As Long
    maxFileTitle As Long
    initialDir As Long
    title As Long
    flags As Long
    fileOffset As Integer
    extensionOffset As Integer
    defaultExtension As Long
    customData As Long
    hook As Long
    templateName As Long
    reserved As Long
    reservedWord As Long
    flagsEx As Long
End Type
Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (ByRef info As OpenFileInfo) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long

' V8i prototype: coordinates in master units, X=north, Y=east.
' Consecutive equal line codes form one open line string.
Public Sub Import()
    Dim folder As String, path As String, lib As String
    Dim row As String, parts As Variant, rows As Variant, separator As String
    Dim n As Long, count As Long, i As Long, j As Long, last As Long
    Dim codes() As String, pts() As Point3d, vertices() As Point3d
    Dim pending As New Collection, added As New Collection
    Dim el As Element, cell As CellElement, lev As Level
    Dim cellCount As Long, lineCount As Long, message As String
    On Error GoTo Failed
    path = PickFile("Vali punktifail", "Punktifailid (*.txt;*.csv)", "*.txt;*.csv", "")
    If Len(path) = 0 Then Exit Sub
    folder = Left$(path, InStrRev(path, "\"))
    lib = folder & "parnu_tm_mkm.cel"
    If Len(Dir$(lib)) = 0 Then
        lib = PickFile("Vali celliteek (SPYMAR)", "Celliteegid (*.cel)", "*.cel", folder)
        If Len(lib) = 0 Then Exit Sub
    End If
    If Not ActiveModelReference.Is3D Then Err.Raise vbObjectError + 3, , "Ava 3D DGN, et sailitada Z-korgused."

    rows = Split(Replace(Replace(ReadText(path), vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For n = 1 To UBound(rows) + 1
        row = Trim$(rows(n - 1))
        If LCase$(row) = "selgitus:" And count > 0 Then Exit For
        If LCase$(Left$(row, 4)) = "sep=" And count = 0 Then
            separator = Mid$(row, 5)
            If separator <> "," And separator <> ";" And separator <> vbTab Then Err.Raise vbObjectError + 11, , "Toetamata eraldaja: " & separator
            row = ""
        End If
        If Len(row) > 0 Then
            If Len(separator) = 0 Then
                separator = ","
                If InStr(row, ";") > 0 Then separator = ";"
                If InStr(row, vbTab) > 0 Then separator = vbTab
            End If
            parts = ParseRow(row, separator)
            If UBound(parts) <> 4 Then Err.Raise vbObjectError + 4, , "Fail: " & path & vbCrLf & "Real " & n & " oodati 5 valja, leiti " & (UBound(parts) + 1) & "." & vbCrLf & Left$(row, 160)
            If count = 0 And LCase$(Join(parts, ",")) = "pnr,x,y,z,kood" Then GoTo NextRow
            If Len(Trim$(parts(0))) = 0 Then Err.Raise vbObjectError + 5, , "Punktinumber puudub real " & n
            Select Case Trim$(parts(4))
                Case "2", "6", "10"
                Case Else: Err.Raise vbObjectError + 6, , "Tundmatu kood real " & n & ": " & parts(4)
            End Select
            count = count + 1
            ReDim Preserve codes(1 To count)
            ReDim Preserve pts(1 To count)
            codes(count) = Trim$(parts(4))
            pts(count) = Point3dFromXYZ(Number(parts(2), n), Number(parts(1), n), Number(parts(3), n))
        End If
NextRow:
    Next n
    If count = 0 Then Err.Raise vbObjectError + 7, , "Punktifail on tuhi."

    ' Build every element before adding any geometry to the model.
    AttachCellLibrary lib
    i = 1
    Do While i <= count
        If codes(i) = "2" Then
            Set cell = CreateCellElement2("SPYMAR", pts(i), Point3dFromXYZ(1, 1, 1), True, Matrix3dIdentity)
            pending.Add cell
            cellCount = cellCount + 1
            i = i + 1
        Else
            last = i
            Do While last < count
                If codes(last + 1) <> codes(i) Then Exit Do
                last = last + 1
            Loop
            If last = i Then Err.Raise vbObjectError + 8, , "Joonekoodil " & codes(i) & " on ainult uks punkt (kirje " & i & ")."
            ReDim vertices(0 To last - i)
            For j = i To last
                vertices(j - i) = pts(j)
                If j > i Then
                    If pts(j).X = pts(j - 1).X And pts(j).Y = pts(j - 1).Y And pts(j).Z = pts(j - 1).Z Then
                        Err.Raise vbObjectError + 9, , "Kattuvad jarjestikused joonepunktid: kirje " & j
                    End If
                End If
            Next j
            Set el = CreateLineElement1(Nothing, vertices)
            pending.Add el
            lineCount = lineCount + 1
            i = last + 1
        End If
    Loop

    For Each el In pending
        If el.IsCellElement Then
            Set lev = EnsureLevel("ALUSVORK")
        Else
            Set lev = EnsureLevel("HOONE")
        End If
        Set el.Level = lev
        ActiveModelReference.AddElement el
        added.Add el
        ApplyLevel el, lev
        el.Redraw msdDrawingModeNormal
    Next el
    MsgBox "Valmis: " & cellCount & " cell, " & lineCount & " joont." & vbCrLf & "Korduv import lisab samad elemendid uuesti.", vbInformation, "ValiImport"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    ' Remove only elements inserted by this invocation.
    For i = added.count To 1 Step -1
        Set el = added(i)
        ActiveModelReference.RemoveElement el
    Next i
    MsgBox "Import katkestati: " & message, vbExclamation, "ValiImport"
End Sub

Private Function Number(ByVal value As String, ByVal row As Long) As Double
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([eE][+-]?[0-9]+)?$"
    value = Replace(Trim$(value), ",", ".")
    If Not re.Test(value) Then Err.Raise vbObjectError + 10, , "Vigane arv real " & row & ": " & value
    Number = Val(value)
End Function

Private Function EnsureLevel(ByVal name As String) As Level
    Dim lev As Level
    On Error Resume Next
    Set lev = ActiveDesignFile.Levels(name)
    On Error GoTo 0
    If lev Is Nothing Then
        Set lev = ActiveDesignFile.AddNewLevel(name)
        ActiveDesignFile.Levels.Rewrite
    End If
    Set EnsureLevel = lev
End Function

Private Sub ApplyLevel(ByVal el As Element, ByVal lev As Level)
    Dim children As ElementEnumerator, child As Element
    Set el.Level = lev
    el.Rewrite
    If el.IsComplexElement Then
        Set children = el.AsComplexElement.GetSubElements
        Do While children.MoveNext
            Set child = children.Current
            ApplyLevel child, lev
        Loop
    End If
End Sub

Private Function PickFile(ByVal title As String, ByVal label As String, ByVal pattern As String, ByVal folder As String) As String
    Dim info As OpenFileInfo, buffer As String, filter As String, result As Long
    buffer = String$(32768, vbNullChar)
    filter = label & vbNullChar & pattern & vbNullChar & vbNullChar
    info.size = LenB(info)
    info.owner = GetActiveWindow()
    info.filter = StrPtr(filter)
    info.filterIndex = 1
    info.file = StrPtr(buffer)
    info.maxFile = Len(buffer)
    info.title = StrPtr(title)
    If Len(folder) > 0 Then info.initialDir = StrPtr(folder)
    info.flags = &H80000 Or &H1000 Or &H800 Or &H8
    If GetOpenFileNameW(info) <> 0 Then
        PickFile = Left$(buffer, InStr(buffer, vbNullChar) - 1)
    Else
        result = CommDlgExtendedError()
        If result <> 0 Then Err.Raise vbObjectError + 12, , "Failidialoogi viga: " & result
    End If
End Function

Private Function ReadText(ByVal path As String) As String
    Dim stream As Object, bytes As Variant, charset As String, content As String
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile path
    charset = "utf-8"
    If stream.size >= 2 Then
        bytes = stream.Read(2)
        If bytes(0) = 255 And bytes(1) = 254 Then charset = "unicode"
        If bytes(0) = 254 And bytes(1) = 255 Then charset = "unicodeFFFE"
    End If
    stream.Position = 0
    stream.Type = 2
    stream.charset = charset
    content = stream.ReadText
    stream.Close
    If Len(content) > 0 Then
        If Left$(content, 1) = ChrW$(-257) Then content = Mid$(content, 2)
    End If
    ReadText = content
End Function

Private Function ParseRow(ByVal row As String, ByVal separator As String) As Variant
    Dim fields() As String, value As String, quoted As Boolean
    Dim i As Long, index As Long, ch As String
    ReDim fields(0 To 0)
    i = 1
    Do While i <= Len(row)
        ch = Mid$(row, i, 1)
        If ch = Chr$(34) Then
            If quoted And Mid$(row, i + 1, 1) = Chr$(34) Then
                value = value & Chr$(34)
                i = i + 1
            Else
                quoted = Not quoted
            End If
        ElseIf ch = separator And Not quoted Then
            fields(index) = Trim$(value)
            index = index + 1
            ReDim Preserve fields(0 To index)
            value = ""
        Else
            value = value & ch
        End If
        i = i + 1
    Loop
    If quoted Then Err.Raise vbObjectError + 13, , "Sulgemata jutumark: " & Left$(row, 160)
    fields(index) = Trim$(value)
    ParseRow = fields
End Function
