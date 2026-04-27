VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentRowResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Resolver")
Option Explicit

Public Function Resolve(ByRef RawRows() As String, ByVal MetaMap As Scripting.Dictionary) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim DateColumn As Long
    DateColumn = GetDateColumn(RawRows)
    Dim TransferColumn As Long
    TransferColumn = GetTransferColumn(RawRows)
    Dim RemarksColumn As Long
    RemarksColumn = GetRemarksColumn(RawRows)
    Dim R As Long
    For R = LBound(RawRows) + 1 To UBound(RawRows)
        Dim RawDate As String
        RawDate = RawRows(R, DateColumn)
        Dim Transfer As String
        Transfer = RawRows(R, TransferColumn)
        Dim Remarks As String
        Remarks = RawRows(R, RemarksColumn)
        Dim C As Long
        For C = LBound(RawRows, 2) To UBound(RawRows, 2)
            If MetaMap.Exists(C) Then
                Dim Meta As App_EnrollmentColumnMeta
                Set Meta = MetaMap.Item(C)
                Dim Row As App_EnrollmentRow
                Set Row = CreateRow(RawDate, Transfer, Remarks, Meta, RawRows(R, C))
                Result.Add Row
            End If
        Next
    Next
    Set Resolve = Result
End Function

Private Function CreateRow(ByVal RawDate As String, ByVal Transfer As String, ByVal Remarks As String, ByVal Meta As App_EnrollmentColumnMeta, ByVal RawText As String) As App_EnrollmentRow
    Dim Result As App_EnrollmentRow
    Set Result = New App_EnrollmentRow
    With Result
        .NormDate = App_DateUtility.NormalizeToDate(RawDate)
        .TypeCode = Meta.TypeCode
        .Grade = Meta.Grade
        .ClassNo = Meta.ClassNo
        .Gender = Meta.Gender
        .Value = App_NumericUtility.ExtractNumber(RawText)
        .Transfer = App_BooleanUtility.NormalizeToBoolean(Transfer)
        .Remarks = Remarks
    End With
    Set CreateRow = Result
End Function

Private Function GetDateColumn(ByRef RawRows() As String) As Long
    Dim C As Long
    For C = LBound(RawRows, 2) To UBound(RawRows, 2)
        If RawRows(1, C) = jpnHizuke Then GetDateColumn = C: Exit Function
    Next
End Function

Private Function GetTransferColumn(ByRef RawRows() As String) As Long
    Dim C As Long
    For C = LBound(RawRows, 2) To UBound(RawRows, 2)
        If RawRows(1, C) = jpnIdou Then GetTransferColumn = C: Exit Function
    Next
End Function

Private Function GetRemarksColumn(ByRef RawRows() As String) As Long
    Dim C As Long
    For C = LBound(RawRows, 2) To UBound(RawRows, 2)
        If RawRows(1, C) = jpnBikou Then GetRemarksColumn = C: Exit Function
    Next
End Function
