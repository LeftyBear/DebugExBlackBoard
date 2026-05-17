VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourRowResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function Resolve(ByVal Records As VBA.Collection, ByVal MetaMap As Scripting.Dictionary) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim R As Long
    For R = 2 To Records.Count
        Dim Rows As VBA.Collection
        Set Rows = Records.Item(R)
        Dim C As Long
        For C = 1 To Rows.Count
            If MetaMap.Exists(C) Then
                Dim Meta As App_ClassHourColumn
                Set Meta = MetaMap.Item(C)
                Dim Row As App_ClassHourRow
                Set Row = CreateRow(Meta, Rows.Item(C))
                Result.Add Row
            End If
        Next
    Next
    Set Resolve = Result
End Function

Private Function CreateRow(ByVal Meta As App_ClassHourColumn, ByVal RawText As String) As App_ClassHourRow
    Dim Row As App_ClassHourRow
    Set Row = New App_ClassHourRow
    If Meta.RawDate <> vbNullString Then
        Row.NormDate = App_DateUtility.NormalizeToDate(RawText)
    End If
    If Meta.TypeCode <> 0 Then
        Row.TypeCode = Meta.TypeCode
        Row.Subject = Meta.Subject
        Row.Period = Meta.Period
        Row.Grade = Meta.Grade
        Row.Value = Inf_NumericUtility.ExtractNumber(RawText)
    End If
    Set CreateRow = Row
End Function
