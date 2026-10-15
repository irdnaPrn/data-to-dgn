Attribute VB_Name = "ElementAndmed"
Option Explicit

' One header, one code, then one original point number per vertex.
' This application name is separate from the RING02 height-text linkage.
Private Const DATA_APP As String = "DATA_TO_DGN_POINTS_V1"

Public Sub Store(ByVal el As Element, ByVal kind As String, ByVal code As String, ByVal numbers As Variant)
    Dim data() As XDatum, i As Long, n As Long
    n = UBound(numbers) - LBound(numbers) + 1
    If n < 1 Then Err.Raise vbObjectError + 100, , "Punktinumbrid puuduvad."
    ReDim data(0 To n + 1)
    SetString data(0), kind
    SetString data(1), code
    For i = 0 To n - 1
        SetString data(i + 2), CStr(numbers(LBound(numbers) + i))
    Next i
    el.SetXData DATA_APP, data
End Sub

Private Sub SetString(ByRef item As XDatum, ByVal value As String)
    ' Reject an oversized value instead of silently truncating the source number.
    If Len(value) > 200 Then Err.Raise vbObjectError + 101, , "Punktinumber voi kood on liiga pikk (ule 200 margi)."
    item.Type = msdXDatumTypeString
    item.Value = value
End Sub

Public Function ReadData(ByVal el As Element, ByRef kind As String, ByRef code As String, ByRef numbers As Variant) As Boolean
    Dim data() As XDatum, values() As String, first As Long, i As Long, n As Long
    kind = "": code = "": numbers = Empty
    If Not el.HasXData(DATA_APP) Then Exit Function
    data = el.GetXData(DATA_APP)
    first = LBound(data)
    n = UBound(data) - first - 1
    If n < 1 Then Err.Raise vbObjectError + 102, , "Elemendi punktiseos on vigane: " & DLongToString(el.ID)
    kind = CStr(data(first).Value)
    code = CStr(data(first + 1).Value)
    ReDim values(0 To n - 1)
    For i = 0 To n - 1
        values(i) = CStr(data(first + i + 2).Value)
    Next i
    numbers = values
    ReadData = True
End Function

Public Function HasData(ByVal el As Element) As Boolean
    HasData = el.HasXData(DATA_APP)
End Function

Public Function NewNumbers(ByVal count As Long, ByVal used As Object) As Variant
    Dim result() As String, i As Long, nextNumber As Double, value As String
    ReDim result(0 To count - 1)
    nextNumber = 1
    For i = 0 To count - 1
        Do
            value = Format$(nextNumber, "0")
            nextNumber = nextNumber + 1
        Loop While used.Exists(value)
        result(i) = value
        used.Add value, True
    Next i
    NewNumbers = result
End Function

Public Function UsedNumbers() As Object
    Dim used As Object, scan As ElementEnumerator, el As Element
    Dim kind As String, code As String, numbers As Variant, i As Long
    Set used = CreateObject("Scripting.Dictionary")
    Set scan = ActiveModelReference.Scan
    Do While scan.MoveNext
        Set el = scan.Current
        If Not el.IsComponentElement And el.IsGraphical Then
            If ReadData(el, kind, code, numbers) Then
                For i = LBound(numbers) To UBound(numbers)
                    used(CStr(numbers(i))) = True
                Next i
            End If
        End If
    Loop
    Set UsedNumbers = used
End Function

Public Function JoinNumbers(ByVal firstNumbers As Variant, ByVal secondNumbers As Variant, _
    ByVal firstForward As Boolean, ByVal secondForward As Boolean, ByVal sharedEnd As Boolean) As Variant
    Dim result() As String, i As Long, j As Long, start As Long, finish As Long, stepValue As Long
    Dim n As Long
    n = UBound(firstNumbers) - LBound(firstNumbers) + UBound(secondNumbers) - LBound(secondNumbers) + 2
    If sharedEnd Then n = n - 1
    ReDim result(0 To n - 1)
    If firstForward Then
        start = LBound(firstNumbers): finish = UBound(firstNumbers): stepValue = 1
    Else
        start = UBound(firstNumbers): finish = LBound(firstNumbers): stepValue = -1
    End If
    For i = start To finish Step stepValue
        result(j) = CStr(firstNumbers(i)): j = j + 1
    Next i
    If secondForward Then
        start = LBound(secondNumbers): finish = UBound(secondNumbers): stepValue = 1
    Else
        start = UBound(secondNumbers): finish = LBound(secondNumbers): stepValue = -1
    End If
    ' A coincident endpoint keeps the first selected line's point number.
    If sharedEnd Then start = start + stepValue
    For i = start To finish Step stepValue
        result(j) = CStr(secondNumbers(i)): j = j + 1
    Next i
    JoinNumbers = result
End Function
