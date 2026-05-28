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

Private Const COLUMN_SUBJECT_NAME   As String = "ã≥â»ñº"
Private Const COLUMN_TARGET_NAME    As String = "ëŒè€äwîN"
Private Const COLUMN_MARK           As String = "ãLçÜ"

Public Function Create(ByVal ColumnName As String, ByVal RawText As String) As Inf_SubjectRow
    Dim Result As Inf_SubjectRow
    Set Result = New Inf_SubjectRow
    If ColumnName = COLUMN_SUBJECT_NAME Then
        Result.Name = RawText
    ElseIf ColumnName = COLUMN_TARGET_NAME Then
        Result.TargetGrade = RawText
    ElseIf ColumnName = COLUMN_MARK Then
        Result.Mark = RawText
    End If
    Set Create = Result
End Function
