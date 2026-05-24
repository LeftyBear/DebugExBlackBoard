VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectRowFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Header As String, ByVal RawText As String) As Inf_SubjectRow
    Dim Result As Inf_SubjectRow
    Set Result = New Inf_SubjectRow
    If Header = jpnKyoukaMei Then
        Result.Name = RawText
    ElseIf Header = jpnTaisyoGakunen Then
        Result.TargetGrade = RawText
    ElseIf Header = jpnKigou Then
        Result.Mark = RawText
    End If
    Set Create = Result
End Function
