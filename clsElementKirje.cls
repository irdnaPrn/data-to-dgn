VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsElementKirje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Source As Element
Public Kind As String
Public Code As String
Public Numbers As Variant
Public PointCount As Long
Public MatchInfo As String
Public NeedsWrite As Boolean
Public Issue As String
Public LevelName As String
Public ElementName As String

Public Function GroupKey() As String
    If Len(Issue) > 0 Then
        GroupKey = "!:" & DLongToString(Source.ID)
    ElseIf Len(Code) > 0 Then
        GroupKey = "K:" & Code
    Else
        GroupKey = "?:" & Kind & vbTab & LevelName & vbTab & ElementName
    End If
End Function
