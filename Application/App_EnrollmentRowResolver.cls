VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentRowResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function Resolve(ByRef RawRows() As String, ByVal MetaMap As Scripting.Dictionary) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim R As Long
    For R = LBound(RawRows) + 1 To UBound(RawRows)
        Dim C As Long
        For C = LBound(RawRows, 2) To UBound(RawRows, 2)
            If MetaMap.Exists(C) Then
                Dim Meta As App_EnrollmentColumnMeta
                Set Meta = MetaMap.Item(C)
                Dim Row As App_EnrollmentRow
                Set Row = CreateRow(Meta, RawRows(R, C))
                Result.Add Row
            End If
        Next
    Next
    Set Resolve = Result
End Function

Private Function CreateRow(ByVal Meta As App_EnrollmentColumnMeta, ByVal RawText As String) As App_EnrollmentRow
    Dim Row As App_EnrollmentRow
    Set Row = New App_EnrollmentRow
    If Meta.ColDate <> vbNullString Then
        Row.NormDate = App_DateUtility.NormalizeToDate(RawText)
    End If
    If Meta.TypeCode <> 0 Then
        Row.TypeCode = Meta.TypeCode
        Row.Grade = Meta.Grade
        Row.ClassNo = Meta.ClassNo
        Row.Gender = Meta.Gender
        Row.Value = App_NumericUtility.ExtractNumber(RawText)
    End If
    If Meta.ColTransfer <> vbNullString Then
        Row.RawTransfer = RawText
    End If
    If Meta.ColRemarks <> vbNullString Then
        Row.RawRemarks = RawText
    End If
    Set CreateRow = Row
End Function
