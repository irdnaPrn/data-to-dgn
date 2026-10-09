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
Private Declare Function SetWindowsHookExW Lib "user32.dll" (ByVal hookType As Long, ByVal callback As Long, ByVal instance As Long, ByVal threadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hook As Long, ByVal code As Long, ByVal windowHandle As Long, ByVal info As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Private Declare Function GetDlgItem Lib "user32.dll" (ByVal dialogHandle As Long, ByVal controlId As Long) As Long
Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal windowHandle As Long, ByVal buffer As Long, ByVal maxCount As Long) As Long
Private Declare Function SetDlgItemTextW Lib "user32.dll" (ByVal dialogHandle As Long, ByVal controlId As Long, ByVal caption As Long) As Long
Private previewHook As Long
Private previewTitle As String, previewButton As String

' V8i prototype: coordinates in master units, X=north, Y=east.
' Consecutive equal line codes form one open line string.
Public Sub Import()
    Dim folder As String, path As String, lib As String, tablePath As String
    Dim row As String, parts As Variant, rows As Variant, separator As String
    Dim n As Long, count As Long, i As Long, j As Long, last As Long
    Dim codes() As String, pointNumbers() As String, pts() As Point3d, vertices() As Point3d
    Dim pending As New Collection, pendingLevels As New Collection, added As New Collection
    Dim ringTexts As New Collection, ringCells As New Collection
    Dim el As Element, cell As CellElement, pointCircle As EllipseElement, textEl As TextElement
    Dim heightText As TextElement
    Dim lev As Level, lineStyle As lineStyle, textStyle As textStyle, heightStyle As textStyle
    Dim engineeringFont As Font
    Dim textPoint As Point3d, heightPoint As Point3d
    Dim heightValue As String
    Dim message As String, definition As Variant
    Dim codeTable As Object, codeCounts As Object
    Dim firstRowChecked As Boolean, separatorSpecified As Boolean
    On Error GoTo Failed
    path = PickFile("Vali punktifail", "Punktifailid (*.txt;*.csv)", "*.txt;*.csv", "")
    If Len(path) = 0 Then Exit Sub
    folder = Left$(path, InStrRev(path, "\"))
    tablePath = folder & "kooditabel.csv"
    If Len(Dir$(tablePath)) = 0 Then
        tablePath = PickFile("Vali kooditabel", "Kooditabel (*.csv;*.txt)", "*.csv;*.txt", folder)
        If Len(tablePath) = 0 Then Exit Sub
    End If
    Set codeTable = ReadCodeTable(tablePath)
    Set codeCounts = CreateObject("Scripting.Dictionary")
    codeCounts.CompareMode = 1
    lib = folder & "parnu_tm_mkm.cel"
    If Len(Dir$(lib)) = 0 Then
        lib = PickFile("Vali celliteek", "Celliteegid (*.cel)", "*.cel", folder)
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
            separatorSpecified = True
            row = ""
        End If
        If Len(row) > 0 Then
            If Len(separator) = 0 Then
                separator = ","
                If InStr(row, ";") > 0 Then separator = ";"
                If InStr(row, vbTab) > 0 Then separator = vbTab
            End If
            parts = ParseRow(row, separator)
            If UBound(parts) <> 4 Then
                If Not firstRowChecked Then
                    firstRowChecked = True
                    If Not separatorSpecified Then separator = ""
                    GoTo NextRow
                End If
                Err.Raise vbObjectError + 4, , "Fail: " & path & vbCrLf & "Real " & n & " oodati 5 valja, leiti " & (UBound(parts) + 1) & "." & vbCrLf & Left$(row, 160)
            End If
            firstRowChecked = True
            If count = 0 And LCase$(Join(parts, ",")) = "pnr,x,y,z,kood" Then GoTo NextRow
            If Len(Trim$(parts(0))) = 0 Then Err.Raise vbObjectError + 5, , "Punktinumber puudub real " & n
            count = count + 1
            ReDim Preserve codes(1 To count)
            ReDim Preserve pointNumbers(1 To count)
            ReDim Preserve pts(1 To count)
            codes(count) = Trim$(parts(4))
            pointNumbers(count) = Trim$(parts(0))
            pts(count) = Point3dFromXYZ(Number(parts(2), n), Number(parts(1), n), Number(parts(3), n))
            If codeCounts.Exists(codes(count)) Then
                codeCounts(codes(count)) = CLng(codeCounts(codes(count))) + 1
            Else
                codeCounts.Add codes(count), 1&
            End If
        End If
NextRow:
    Next n
    If count = 0 Then Err.Raise vbObjectError + 7, , "Punktifail on tuhi."
    If Not ConfirmImport(codeCounts, codeTable, count) Then Exit Sub

    ' Build every element before adding any geometry to the model.
    AttachCellLibrary lib
    i = 1
    Do While i <= count
        If Not codeTable.Exists(codes(i)) Then
            i = i + 1
        Else
            definition = codeTable(codes(i))
            If definition(0) = "CELL" Then
                Set cell = CreateCellElement2(definition(2), pts(i), Point3dFromXYZ(CDbl(definition(3)), CDbl(definition(3)), 1), True, Matrix3dIdentity)
                pending.Add cell
                pendingLevels.Add definition(1)
                If UCase$(CStr(definition(2))) = "RING02" Then
                    If engineeringFont Is Nothing Then Set engineeringFont = FindFont("ENGINEERING")
                    heightValue = Replace(Format$(pts(i).Z, "0.00"), ",", ".")
                    heightPoint = Point3dFromXYZ(pts(i).X + 0.3, pts(i).Y - 0.5, pts(i).Z)
                    Set heightText = CreateTextElement1(Nothing, heightValue, heightPoint, Matrix3dIdentity)
                    Set heightStyle = heightText.textStyle
                    heightStyle.Height = 0.85
                    heightStyle.Width = 0.65
                    heightStyle.Justification = msdTextJustificationLeftBottom
                    Set heightStyle.Font = engineeringFont
                    Set heightText.textStyle = heightStyle
                    ringTexts.Add heightText
                    ringCells.Add cell
                    pending.Add heightText
                    pendingLevels.Add definition(1)
                End If
                i = i + 1
            Else
                last = i
                Do While last < count
                    If codes(last + 1) <> codes(i) Then Exit Do
                    last = last + 1
                Loop
                If last = i Then
                    i = i + 1
                Else
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
                    Set lineStyle = FindLineStyle(CStr(definition(2)))
                    Set el.lineStyle = lineStyle
                    pending.Add el
                    pendingLevels.Add definition(1)
                    i = last + 1
                End If
            End If
        End If
    Loop

    For i = 1 To count
        Set pointCircle = CreateEllipseElement2(Nothing, pts(i), 0.08, 0.08, Matrix3dIdentity, msdFillModeNotFilled)
        pending.Add pointCircle
        pendingLevels.Add "MOOTMPUNKT"

        textPoint = Point3dFromXYZ(pts(i).X + 0.25, pts(i).Y + 0.15, pts(i).Z)
        Set textEl = CreateTextElement1(Nothing, pointNumbers(i), textPoint, Matrix3dIdentity)
        Set textStyle = textEl.textStyle
        textStyle.Height = 0.2
        textStyle.Width = 0.2
        textStyle.Justification = msdTextJustificationLeftCenter
        Set textEl.textStyle = textStyle
        pending.Add textEl
        pendingLevels.Add "MOOTNR"

        heightValue = Replace(Format$(pts(i).Z, "0.00"), ",", ".")
        heightPoint = Point3dFromXYZ(textPoint.X, textPoint.Y - 0.3, textPoint.Z)
        Set heightText = CreateTextElement1(Nothing, heightValue, heightPoint, Matrix3dIdentity)
        Set heightText.textStyle = textStyle
        pending.Add heightText
        pendingLevels.Add "MOOTKORG-EH2000"
    Next i

    For i = 1 To pending.count
        Set el = pending(i)
        Set lev = EnsureLevel(CStr(pendingLevels(i)))
        Set el.Level = lev
        ActiveModelReference.AddElement el
        added.Add el
        ApplyLevel el, lev
        el.Redraw msdDrawingModeNormal
    Next i
    For i = 1 To ringTexts.count
        KorgusePooramine.LinkTextToCell ringTexts(i), ringCells(i)
    Next i
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

Private Function ConfirmImport(ByVal codeCounts As Object, ByVal codeTable As Object, ByVal pointCount As Long) As Boolean
    Dim pages As New Collection
    Dim codeKey As Variant, codeLabel As String, entry As String, pageText As String
    Dim pageIndex As Long, pageRows As Long, prompt As String, buttonCaption As String

    ' Keep each dialog short enough to display the complete code list.
    For Each codeKey In codeCounts.Keys
        codeLabel = CStr(codeKey)
        If Len(codeLabel) = 0 Then codeLabel = "(tuhi)"
        entry = codeLabel & " | " & CStr(codeCounts(codeKey))
        If Not codeTable.Exists(codeKey) Then entry = entry & " | kood puudub"
        If Len(pageText) > 0 Then
            If Len(pageText) + Len(entry) + 2 > 750 Or pageRows >= 20 Then
                pages.Add pageText
                pageText = ""
                pageRows = 0
            End If
        End If
        If Len(pageText) > 0 Then pageText = pageText & vbCrLf
        pageText = pageText & entry
        pageRows = pageRows + 1
    Next codeKey
    If Len(pageText) > 0 Then pages.Add pageText

    For pageIndex = 1 To pages.count
        If pageIndex = pages.count Then
            buttonCaption = "Import"
        Else
            buttonCaption = "Edasi"
        End If
        prompt = "Punkte kokku: " & pointCount & vbCrLf & "Kood | Punkte | Info" & vbCrLf & vbCrLf & _
            CStr(pages(pageIndex))
        If PreviewMessageBox(prompt, "Impordi eelvaade (" & pageIndex & "/" & pages.count & ")", buttonCaption) <> vbOK Then Exit Function
    Next pageIndex
    ConfirmImport = True
End Function

Private Function PreviewMessageBox(ByVal prompt As String, ByVal title As String, ByVal buttonCaption As String) As Long
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    previewTitle = title
    previewButton = buttonCaption
    ' Thread-local CBT hook: rename this existing MsgBox before it is shown.
    previewHook = SetWindowsHookExW(5, AddressOf PreviewDialogHook, 0, GetCurrentThreadId())
    If previewHook = 0 Then Err.Raise vbObjectError + 27, , "Impordidialoogi nuppude seadistamine ebaonnestus."
    PreviewMessageBox = MsgBox(prompt, vbOKCancel Or vbInformation, title)
    If previewHook <> 0 Then UnhookWindowsHookEx previewHook
    previewHook = 0
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If previewHook <> 0 Then UnhookWindowsHookEx previewHook
    previewHook = 0
    Err.Raise errorNumber, "ValiImport.PreviewMessageBox", errorText
End Function

Public Function PreviewDialogHook(ByVal code As Long, ByVal windowHandle As Long, ByVal info As Long) As Long
    Dim titleBuffer As String, titleLength As Long, cancelCaption As String
    Dim hookHandle As Long, renamed As Boolean
    ' Errors must never propagate across a Windows callback boundary.
    On Error Resume Next
    hookHandle = previewHook
    If code = 5 Then ' HCBT_ACTIVATE
        titleBuffer = String$(256, vbNullChar)
        titleLength = GetWindowTextW(windowHandle, StrPtr(titleBuffer), Len(titleBuffer))
        If Left$(titleBuffer, titleLength) = previewTitle Then
            If GetDlgItem(windowHandle, vbOK) <> 0 And GetDlgItem(windowHandle, vbCancel) <> 0 Then
                cancelCaption = "Katkesta"
                SetDlgItemTextW windowHandle, vbOK, StrPtr(previewButton)
                SetDlgItemTextW windowHandle, vbCancel, StrPtr(cancelCaption)
                renamed = True
            End If
        End If
    End If
    PreviewDialogHook = CallNextHookEx(hookHandle, code, windowHandle, info)
    If renamed Then
        UnhookWindowsHookEx hookHandle
        previewHook = 0
    End If
End Function

Private Function FindFont(ByVal name As String) As Font
    Dim fontItem As Font
    On Error Resume Next
    Set fontItem = ActiveDesignFile.Fonts.Item(name)
    On Error GoTo 0
    If fontItem Is Nothing Then Err.Raise vbObjectError + 26, , "Fonti ei leitud: " & name
    Set FindFont = fontItem
End Function

Private Function FindLineStyle(ByVal name As String) As lineStyle
    Dim style As lineStyle
    On Error Resume Next
    If IsNumeric(name) Then
        Set style = ActiveDesignFile.LineStyles.Item(CLng(name))
    Else
        Set style = ActiveDesignFile.LineStyles.Item(name)
    End If
    On Error GoTo 0
    If style Is Nothing Then
        ActiveDesignFile.LineStyles.Refresh
        On Error Resume Next
        If IsNumeric(name) Then
            Set style = ActiveDesignFile.LineStyles.Item(CLng(name))
        Else
            Set style = ActiveDesignFile.LineStyles.Item(name)
        End If
        On Error GoTo 0
    End If
    If style Is Nothing Then Err.Raise vbObjectError + 23, , "Joonestiili ei leitud: " & name
    Set FindLineStyle = style
End Function

Private Function ReadCodeTable(ByVal path As String) As Object
    Dim table As Object, rows As Variant, parts As Variant
    Dim row As String, separator As String, code As String, itemType As String
    Dim n As Long, fieldCount As Long, cellScale As Double
    Set table = CreateObject("Scripting.Dictionary")
    table.CompareMode = 1
    rows = Split(Replace(Replace(ReadText(path), vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For n = 1 To UBound(rows) + 1
        row = Trim$(rows(n - 1))
        If Len(row) > 0 Then
            If LCase$(Left$(row, 4)) = "sep=" And table.count = 0 Then
                separator = Mid$(row, 5)
                If separator <> "," And separator <> ";" And separator <> vbTab Then Err.Raise vbObjectError + 22, , "Kooditabelis on toetamata eraldaja: " & separator
                GoTo NextCodeRow
            End If
            If Len(separator) = 0 Then
                separator = ";"
                If InStr(row, ";") = 0 And InStr(row, vbTab) > 0 Then separator = vbTab
                If InStr(row, ";") = 0 And InStr(row, vbTab) = 0 Then separator = ","
            End If
            parts = ParseRow(row, separator)
            fieldCount = UBound(parts) + 1
            If fieldCount <> 6 Then Err.Raise vbObjectError + 14, , "Kooditabeli real " & n & " oodati 6 valja, leiti " & fieldCount & "."
            If table.count = 0 And LCase$(Trim$(parts(0))) = "kood" Then GoTo NextCodeRow
            code = Trim$(parts(0))
            itemType = UCase$(Trim$(parts(2)))
            If Len(code) = 0 Then Err.Raise vbObjectError + 15, , "Kooditabeli real " & n & " puudub kood."
            If table.Exists(code) Then Err.Raise vbObjectError + 16, , "Kooditabelis kordub kood " & code & "."
            If itemType <> "CELL" And itemType <> "JOON" Then Err.Raise vbObjectError + 18, , "Kooditabeli real " & n & " peab tuup olema CELL voi JOON."
            If Len(Trim$(parts(3))) = 0 Then Err.Raise vbObjectError + 19, , "Kooditabeli real " & n & " puudub leveli nimi."
            If Len(Trim$(parts(4))) = 0 Then Err.Raise vbObjectError + 20, , "Kooditabeli real " & n & " puudub joone voi celli nimi."
            cellScale = 1
            If itemType = "CELL" Then
                If Len(Trim$(parts(5))) = 0 Then Err.Raise vbObjectError + 24, , "Kooditabeli real " & n & " puudub celli skaala."
                cellScale = Number(parts(5), n)
                If cellScale <= 0 Then Err.Raise vbObjectError + 25, , "Kooditabeli real " & n & " peab celli skaala olema suurem kui null."
            End If
            table.Add code, Array(itemType, Trim$(parts(3)), Trim$(parts(4)), cellScale)
        End If
NextCodeRow:
    Next n
    If table.count = 0 Then Err.Raise vbObjectError + 21, , "Kooditabel on tuhi."
    Set ReadCodeTable = table
End Function

Private Function Number(ByVal value As String, ByVal row As Long) As Double
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([eE][+-]?[0-9]+)?$"
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
    el.Color = -1 ' ByLevel, including nested cell components.
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
