VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentRowFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Meta As Inf_EnrollmentColumnMeta, ByVal RawText As String) As Inf_EnrollmentRow
    Dim Row As Inf_EnrollmentRow
    Set Row = New Inf_EnrollmentRow
    If Meta.RawDate <> vbNullString Then
        Row.NormDate = Inf_DateUtility.NormalizeToDate(RawText)
    ElseIf Meta.RawID <> vbNullString Then
        Row.RawID = Meta.RawID
        Row.Value = Inf_NumericUtility.ExtractNumber(RawText)
    ElseIf Meta.RawTransfer <> vbNullString Then
        Row.RawTransfer = RawText
    ElseIf Meta.RawRemarks <> vbNullString Then
        Row.Remarks = RawText
    End If
    Set Create = Row
End Function
