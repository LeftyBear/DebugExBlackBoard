VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourPersistence"
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

Public Function ReadAll(ByVal SchoolYear As Long) As Inf_ClassHourRows
    Dim FilePath As String
    FilePath = BuildFilePath(SchoolYear)
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Set ReadAll = Inf_ClassHourRowsFactory.Create(RawRows)
End Function

Public Sub WriteAll(ByVal SchoolYear As Long, ByVal Rows As Inf_ClassHourRows)
    Dim FilePath As String
    FilePath = BuildFilePath(SchoolYear)
    Dim Header() As Variant
    Header = This.CSV.ReadAll(FilePath).GetHeaders
    Dim Body() As Variant
    Body = GroupByDate(Rows)
    Dim RawRows As Inf_RawRows
    Set RawRows = BuildRows(Header, Body)
    This.CSV.WriteAll FilePath, RawRows
End Sub

Public Sub Delete(ByVal SchoolYear As Long, ByVal Key As String)
    Dim FilePath As String
    FilePath = BuildFilePath(SchoolYear)
    This.CSV.Delete FilePath, Key
End Sub

Private Function GroupByDate(ByVal Rows As Inf_ClassHourRows) As Variant()
    Dim UpperColumn As Long
    UpperColumn = GetColumnsCount(Rows)
    Dim UpperRow As Long
    UpperRow = Rows.Count / UpperColumn
    Dim Result() As Variant
    ReDim Result(1 To UpperRow, 1 To UpperColumn)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim R As Long
        R = ((i - 1) \ UpperColumn) + 1
        Dim C As Long
        C = ((i - 1) Mod UpperColumn) + 1
        If 0 < Rows.Item(i).NormDate Then
            Result(R, C) = Rows.Item(i).NormDate
        ElseIf Rows.Item(i).RawID <> vbNullString Then
            Result(R, C) = Rows.Item(i).RawValue
        End If
    Next
    GroupByDate = Result
End Function

Private Function GetColumnsCount(ByVal Rows As Inf_ClassHourRows) As Long
    Dim i As Long
    For i = 1 To Rows.Count
        If 0 < Rows.Item(i).NormDate Then
            Dim IsTarget As Boolean
            If IsTarget = False Then
                IsTarget = True
            Else
                Exit For
            End If
        End If
        If IsTarget Then
            Dim Count As Long
            Count = Count + 1
        End If
    Next
    GetColumnsCount = Count
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

Private Function BuildFilePath(ByVal SchoolYear As Long) As String
    Dim BaseFolderPath As String
    BaseFolderPath = This.Provider.GetBaseFolderPath
    Dim FileName As String
    FileName = "ClassHour" & CStr(SchoolYear) & ".csv"
    BuildFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "data", "schoolyears", CStr(SchoolYear), FileName)
End Function
