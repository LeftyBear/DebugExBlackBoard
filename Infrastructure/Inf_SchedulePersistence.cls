VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchedulePersistence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Persistence")
Option Explicit

Private Type Member
    CSV         As Inf_CSVPersistence
    Provider    As Inf_WorkbookPathProvider
End Type

Private This As Member

Public Sub Inject(ByVal CSV As Inf_CSVPersistence)
    Set This.CSV = CSV
    Set This.Provider = New Inf_WorkbookPathProvider
End Sub

Public Function ReadAll(ByVal SchoolYear As Long) As Inf_ScheduleRows
    Dim FilePath As String
    FilePath = BuildFilePath(SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Exit Function
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Set ReadAll = Inf_ScheduleRowsFactory.Create(RawRows)
End Function

Public Sub WriteAll(ByVal SchoolYear As Long, ByVal Rows As Inf_ScheduleRows)
    Dim FilePath As String
    FilePath = BuildFilePath(SchoolYear)
    Dim Header As VBA.Collection
    Set Header = GetHeader(FilePath)
    Dim Body As VBA.Collection
    Set Body = GroupByDate(Rows)
    Dim RawRows As Inf_RawRows
    Set RawRows = BuildRows(Header, Body)
    This.CSV.WriteAll FilePath, RawRows
End Sub

Public Sub Delete(ByVal SchoolYear As Long, ByVal Key As String)
    Dim FilePath As String
    FilePath = BuildFilePath(SchoolYear)
    This.CSV.Delete FilePath, Key
End Sub

Private Function GetHeader(ByVal FilePath As String) As VBA.Collection
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Set GetHeader = RawRows.GetHeaders
End Function

Private Function GroupByDate(ByVal Rows As Inf_ScheduleRows) As VBA.Collection
    Dim CurrentDate As Date
    CurrentDate = Rows.Item(1).NormDate
    Dim Records As VBA.Collection
    Set Records = New VBA.Collection
    Dim Fields As VBA.Collection
    Set Fields = New VBA.Collection
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_ScheduleRow
        Set Row = Rows.Item(i)
        If Row.NormDate <> vbNullString Then
            Fields.Add CStr(Row.NormDate)
        Else
            Fields.Add Row.Value
        End If
        If CurrentDate <> Row.NormDate Then
            CurrentDate = Row.NormDate
            Records.Add Fields
            Set Fields = New VBA.Collection
        End If
    Next
    Set GroupByDate = Records
End Function

Private Function BuildRows(ByVal Header As VBA.Collection, ByVal Body As VBA.Collection) As Inf_RawRows
    Dim Result As Inf_RawRows
    Set Result = New Inf_RawRows
    Dim C As Long
    For C = 1 To Header.Count
        Result.AddRow Header.Item(C)
    Next
    Result.AddRowsInternal
    Dim R As Long
    For R = 1 To Body.Count
        Result.AddRows Body.Item(R)
    Next
    Set BuildRows = Result
End Function

Private Function BuildFilePath(ByVal SchoolYear As Long) As String
    Dim BaseFolderPath As String
    BaseFolderPath = This.Provider.GetBaseFolderPath
    Dim FileName As String
    FileName = "Schedule" & CStr(SchoolYear) & ".csv"
    BuildFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "data", "schoolyears", CStr(SchoolYear), FileName)
End Function
