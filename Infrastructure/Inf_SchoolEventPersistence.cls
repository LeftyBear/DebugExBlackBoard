VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchoolEventPersistence"
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

Public Function ReadAll() As Inf_SchoolEventRows
    Dim FilePath As String
    FilePath = BuildFilePath
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Dim Factory As New Inf_SchoolEventRowsFactory
    Set ReadAll = Factory.Create(RawRows)
End Function

Public Sub WriteAll(ByVal Rows As Inf_SchoolEventRows)
    Dim Headers As VBA.Collection
    Set Headers = GetHeaders
    Dim Body As VBA.Collection
    Set Body = GroupByRecord(Rows)
    Dim RawRows As Inf_RawRows
    Set RawRows = BuildRows(Headers, Body)
    Dim FilePath As String
    FilePath = BuildFilePath
    This.CSV.WriteAll FilePath, RawRows
End Sub

Public Sub Delete(ByVal SchoolYear As Long, ByVal Key As String)
    Dim FilePath As String
    FilePath = BuildFilePath
    This.CSV.Delete FilePath, Key
End Sub

Public Function GetHeaders() As VBA.Collection
    Dim FilePath As String
    FilePath = BuildFilePath
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Set GetHeaders = RawRows.GetHeaders
End Function

Private Function GroupByRecord(ByVal Rows As Inf_SchoolEventRows) As VBA.Collection
    Dim Records As VBA.Collection
    Set Records = New VBA.Collection
    Dim Fields As VBA.Collection
    Set Fields = New VBA.Collection
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SchoolEventRow
        Set Row = Rows.Item(i)
        Fields.Add Row.Name
        Records.Add Fields
        Set Fields = New VBA.Collection
    Next
    Set GroupByRecord = Records
End Function

Private Function BuildRows(ByVal Headers As VBA.Collection, ByVal Body As VBA.Collection) As Inf_RawRows
    Dim Result As Inf_RawRows
    Set Result = New Inf_RawRows
    Dim C As Long
    For C = 1 To Headers.Count
        Result.AddRow Headers.Item(C)
    Next
    Result.AddRowsInternal
    Dim R As Long
    For R = 1 To Body.Count
        Result.AddRows Body.Item(R)
    Next
    Set BuildRows = Result
End Function

Private Function BuildFilePath() As String
    Dim BaseFolderPath As String
    BaseFolderPath = This.Provider.GetBaseFolderPath
    Dim FileName As String
    FileName = "SchoolEventCatalog.csv"
    BuildFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "Config", FileName)
End Function
