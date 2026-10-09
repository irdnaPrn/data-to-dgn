Attribute VB_Name = "JoonteYhendamine"
Option Explicit

Public Sub Start()
    CommandState.StartPrimitive New clsJoinLines
End Sub

Public Sub JoinPair(ByVal firstLine As Element, ByVal secondLine As Element)
    Dim firstPoints() As Point3d, secondPoints() As Point3d, joinedPoints() As Point3d
    Dim firstEnd As Long, secondEnd As Long, firstStart As Long, secondStart As Long
    Dim a As Long, b As Long, i As Long, nextIndex As Long, stepFirst As Long, stepSecond As Long
    Dim distance As Double, bestDistance As Double
    Dim joined As LineElement, firstBackup As Element, secondBackup As Element
    Dim joinedAdded As Boolean, firstRemoved As Boolean, secondRemoved As Boolean
    Dim errorText As String, rollbackText As String
    On Error GoTo Failed

    firstPoints = firstLine.AsLineElement.GetVertices
    secondPoints = secondLine.AsLineElement.GetVertices
    If UBound(firstPoints) <= LBound(firstPoints) Or UBound(secondPoints) <= LBound(secondPoints) Then
        Err.Raise vbObjectError + 40, , "Molemal joonel peab olema vahemalt kaks tippu."
    End If

    ' Test all four endpoint pairs using XYZ, including elevation.
    firstEnd = UBound(firstPoints)
    secondStart = LBound(secondPoints)
    bestDistance = DistanceSquared(firstPoints(firstEnd), secondPoints(secondStart))
    For a = 0 To 1
        If a = 0 Then i = LBound(firstPoints) Else i = UBound(firstPoints)
        For b = 0 To 1
            If b = 0 Then nextIndex = LBound(secondPoints) Else nextIndex = UBound(secondPoints)
            distance = DistanceSquared(firstPoints(i), secondPoints(nextIndex))
            If distance < bestDistance Then
                bestDistance = distance
                firstEnd = i
                secondStart = nextIndex
            End If
        Next b
    Next a

    If firstEnd = UBound(firstPoints) Then
        firstStart = LBound(firstPoints)
        stepFirst = 1
    Else
        firstStart = UBound(firstPoints)
        stepFirst = -1
    End If
    If secondStart = LBound(secondPoints) Then
        secondEnd = UBound(secondPoints)
        stepSecond = 1
    Else
        secondEnd = LBound(secondPoints)
        stepSecond = -1
    End If

    ReDim joinedPoints(0 To UBound(firstPoints) - LBound(firstPoints) + UBound(secondPoints) - LBound(secondPoints) + 1)
    nextIndex = 0
    For i = firstStart To firstEnd Step stepFirst
        joinedPoints(nextIndex) = firstPoints(i)
        nextIndex = nextIndex + 1
    Next i
    ' A shared endpoint is included once; a gap becomes a straight segment.
    If bestDistance = 0 Then secondStart = secondStart + stepSecond
    For i = secondStart To secondEnd Step stepSecond
        joinedPoints(nextIndex) = secondPoints(i)
        nextIndex = nextIndex + 1
    Next i
    ReDim Preserve joinedPoints(0 To nextIndex - 1)

    ' The first selected element supplies the non-geometric attributes.
    Set joined = CreateLineElement1(firstLine, joinedPoints)
    Set firstBackup = firstLine.Clone
    Set secondBackup = secondLine.Clone
    ActiveModelReference.AddElement joined
    joinedAdded = True
    ActiveModelReference.RemoveElement firstLine
    firstRemoved = True
    ActiveModelReference.RemoveElement secondLine
    secondRemoved = True
    joined.Redraw msdDrawingModeNormal
    Exit Sub

Failed:
    errorText = Err.Description
    ' Restore deleted originals if replacing the pair fails midway.
    On Error Resume Next
    If firstRemoved Then
        Err.Clear
        ActiveModelReference.AddElement firstBackup
        If Err.Number <> 0 Then rollbackText = rollbackText & vbCrLf & "Esimese joone taastamine ebaonnestus: " & Err.Description
    End If
    If secondRemoved Then
        Err.Clear
        ActiveModelReference.AddElement secondBackup
        If Err.Number <> 0 Then rollbackText = rollbackText & vbCrLf & "Teise joone taastamine ebaonnestus: " & Err.Description
    End If
    If joinedAdded Then
        Err.Clear
        ActiveModelReference.RemoveElement joined
        If Err.Number <> 0 Then rollbackText = rollbackText & vbCrLf & "Uue joone eemaldamine ebaonnestus: " & Err.Description
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 41, "JoonteYhendamine.JoinPair", errorText & rollbackText
End Sub

Private Function DistanceSquared(ByRef firstPoint As Point3d, ByRef secondPoint As Point3d) As Double
    DistanceSquared = (firstPoint.X - secondPoint.X) ^ 2 + (firstPoint.Y - secondPoint.Y) ^ 2 + (firstPoint.Z - secondPoint.Z) ^ 2
End Function
