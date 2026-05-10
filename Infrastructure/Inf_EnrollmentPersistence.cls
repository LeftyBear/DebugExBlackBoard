VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentPersistence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Persistence")
Option Explicit

Private Type Member
    CSV As Inf_CSVPersistence
End Type

Private This As Member

Public Sub Initialize(ByVal CSV As Inf_CSVPersistence)
    Set This.CSV = CSV
End Sub

Public Function ReadAll(ByVal SchoolYear As Long) As Inf_EnrollmentRows
    Dim FilePath As String
    FilePath = BuiledFilePath(SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Exit Function
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Dim Factory As Inf_EnrollmentRowsFactory
    Set ReadAll = Factory.Create(RawRows)
End Function

Public Sub WriteAll(ByVal SchoolYear As Long, ByVal Rows As Inf_EnrollmentRows)
    Dim FilePath As String
    FilePath = BuiledFilePath(SchoolYear)
    Dim Header As VBA.Collection
    Set Header = GetHeader(FilePath)
    Dim Body As VBA.Collection
    Set Body = GroupByDate(Rows)
    Dim RawRows As VBA.Collection
    Set RawRows = BuiledRawRows(Header, Body)
    This.CSV.WriteAll FilePath, RawRows
End Sub

Public Sub Delete(ByVal SchoolYear As Long, ByVal Key As String)
    Dim FilePath As String
    FilePath = BuiledFilePath(SchoolYear)
    This.CSV.Delete FilePath, Key
End Sub

Private Function GetHeader(ByVal FilePath As String) As VBA.Collection
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Set GetHeader = RawRows.GetHeaderRows
End Function

Private Function GroupByDate(ByVal Rows As Inf_EnrollmentRows) As VBA.Collection
    Dim Records As VBA.Collection
    Set Records = New VBA.Collection
    Dim CurrentDate As Date
    CurrentDate = Rows.Item(1).NormDate
    Dim Fields As VBA.Collection
    Set Fields = New VBA.Collection
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_EnrollmentRow
        Set Row = Rows.Item(i)
        If Row.NormDate <> vbNullString Then
            Fields.Add CStr(Row.NormDate)
        ElseIf Row.RawID <> vbNullString Then
            Fields.Add CStr(Row.Value)
        ElseIf Row.Transfer <> vbNullString Then
            Fields.Add Row.Transfer
        ElseIf Row.Remarks <> vbNullString Then
            Fields.Add Row.Remarks
        End If
        If CurrentDate <> Row.NormDate Then
            CurrentDate = Row.NormDate
            Records.Add Fields
            Set Fields = New VBA.Collection
        End If
    Next
    Set GroupByDate = Records
End Function

Private Function BuiledRawRows(ByVal Header As VBA.Collection, ByVal Body As VBA.Collection) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Headers As VBA.Collection
    Set Headers = New VBA.Collection
    Dim C As Long
    For C = 1 To Header.Count
        Headers.Add Header.Item(C)
    Next
    Result.Add Headers
    Dim R As Long
    For R = 1 To Body.Count
        Result.Add Body.Item(R)
    Next
    Set BuiledRawRows = Result
End Function

Private Function BuiledFilePath(ByVal SchoolYear As Long) As String
    Dim Provider As Inf_WorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim BaseFolderPath As String
    BaseFolderPath = Provider.GetBaseFolderPath
    Dim FileName As String
    FileName = "Enrollment" & CStr(SchoolYear) & ".csv"
    Dim FilePath As String
    FilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "data", "schoolyears", CStr(SchoolYear), FileName)
    BuiledFilePath = FilePath
End Function
