Attribute VB_Name = "TooriistaAknad"
Option Explicit

' Forms contain only code; their controls are built at runtime.
Public Function Control(ByVal host As Object, ByVal progId As String, ByVal name As String, _
    ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single, _
    Optional ByVal caption As String = "") As Object
    Dim item As Object
    Set item = host.Controls.Add("Forms." & progId & ".1", name, True)
    item.left = left: item.top = top: item.width = width: item.height = height
    If Len(caption) > 0 Then item.caption = caption
    Set Control = item
End Function

Public Function CodeTable() As Object
    Dim path As String, folder As String
    folder = Left$(ActiveDesignFile.FullName, InStrRev(ActiveDesignFile.FullName, "\"))
    path = folder & "kooditabel.csv"
    If Len(Dir$(path)) = 0 Then
        path = ValiImport.PickFile("Vali kooditabel", "Kooditabel (*.csv;*.txt)", "*.csv;*.txt", folder)
    End If
    If Len(path) > 0 Then Set CodeTable = ValiImport.ReadCodeTable(path)
End Function
