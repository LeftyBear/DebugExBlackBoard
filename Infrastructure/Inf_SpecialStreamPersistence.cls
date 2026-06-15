VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SpecialStreamPersistence"
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

Public Function ReadAll() As Inf_SpecialStreamRows
    Dim FilePath As String
    FilePath = BuildFilePath
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Set ReadAll = Inf_SpecialStreamRowsFactory.Create(RawRows)
End Function

Public Sub WriteAll(ByVal Rows As Inf_SpecialStreamRows)
    Dim FilePath As String
    FilePath = BuildFilePath
    Dim Header() As Variant
    Header = This.CSV.ReadAll(FilePath).GetHeader
    Dim Body() As Variant
    Body = BuildBody(Rows)
    Dim RawRows As Inf_RawRows
    Set RawRows = BuildRows(Header, Body)
    This.CSV.WriteAll FilePath, RawRows
End Sub

Public Sub Delete(ByVal Key As String)
    Dim FilePath As String
    FilePath = BuildFilePath
    This.CSV.Delete FilePath, Key
End Sub

Private Function BuildBody(ByVal Rows As Inf_SpecialStreamRows) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To Rows.Count, 1 To 1)
    Dim i As Long
    For i = 1 To Rows.Count
        Result(i, 1) = Rows.Item(i).Name
    Next
    BuildBody = Result
End Function

Private Function BuildRows(ByRef Header() As Variant, ByRef Body() As Variant) As Inf_RawRows
    Dim Result As Inf_RawRows
    Set Result = New Inf_RawRows
    Dim i As Long
    For i = LBound(Header) To UBound(Header)
        Result.AddRow Header(i)
    Next
    Result.AddRowsInternal
    Dim R As Long
    For R = LBound(Body) To UBound(Body)
        Dim C As Long
        For C = LBound(Body, 2) To UBound(Body, 2)
            Result.AddRow Body(R, C)
        Next
        Result.AddRowsInternal
    Next
    Set BuildRows = Result
End Function

Private Function BuildFilePath() As String
    Dim BaseFolderPath As String
    BaseFolderPath = This.Provider.GetBaseFolderPath
    Dim FileName As String
    FileName = "SpecialStreamCatalog.csv"
    BuildFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "Config", FileName)
End Function
