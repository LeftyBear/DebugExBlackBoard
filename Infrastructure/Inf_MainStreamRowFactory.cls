VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_MainStreamRowFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Private Const UPPER_GRADE As String = "ïÅí äwîNêî"
Private Const UPPER_CLASS  As String = "ïÅí äwãâêî"

Public Function Create(ByVal Header As String, ByVal RawText As String) As Inf_MainStreamRow
    Dim Result As Inf_MainStreamRow
    Set Result = New Inf_MainStreamRow
    If Header = UPPER_GRADE Then
        Result.UpperGrade = RawText
    ElseIf Header = UPPER_CLASS Then
        Result.UpperClassNo = RawText
    End If
    Set Create = Result
End Function
