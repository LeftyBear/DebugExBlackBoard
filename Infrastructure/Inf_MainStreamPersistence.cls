VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_MainStreamPersistence"
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

Public Function ReadAll() As Inf_MainStreamRows
    Dim FilePath As String
    FilePath = BuildFilePath
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Set ReadAll = Inf_MainStreamRowsFactory.Create(RawRows)
End Function

Public Sub WriteAll(ByVal Rows As Inf_MainStreamRows)
    Dim Headers As VBA.Collection
    Set Headers = GetHeaders
    Dim Body As VBA.Collection
    Set Body = GetValues(Rows)
    Dim RawRows As Inf_RawRows
    Set RawRows = BuildRows(Headers, Body)
    Dim FilePath As String
    FilePath = BuildFilePath
    This.CSV.WriteAll FilePath, RawRows
End Sub

Public Sub Delete(ByVal Key As String)
    Dim FilePath As String
    FilePath = BuildFilePath
    This.CSV.Delete FilePath, Key
End Sub

Public Function GetHeaders() As Variant
    Dim FilePath As String
    FilePath = BuildFilePath
    GetHeaders = This.CSV.ReadAll(FilePath).GetHeaders
End Function

Private Function GetValues(ByVal Rows As Inf_MainStreamRows) As VBA.Collection
    Dim Records As VBA.Collection
    Set Records = New VBA.Collection
    Dim Fields As VBA.Collection
    Set Fields = New VBA.Collection
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_MainStreamRow
        Set Row = Rows.Item(i)
        If Row.UpperGrade <> vbNullString Then
            Fields.Add CStr(Row.UpperGrade)
        ElseIf Row.UpperClassNo <> vbNullString Then
            Fields.Add CStr(Row.UpperClassNo)
        End If
    Next
    Records.Add Fields
    Set GetValues = Records
End Function

Private Function BuildRows(ByVal Headers As VBA.Collection, ByVal Values As VBA.Collection) As Inf_RawRows
    Dim Result As Inf_RawRows
    Set Result = New Inf_RawRows
    Dim C As Long
    For C = 1 To Headers.Count
        Result.AddRow Headers.Item(C)
    Next
    Result.AddRowsInternal
    Dim R As Long
    For R = 1 To Values.Count
        Result.AddRows Values.Item(R)
    Next
    Set BuildRows = Result
End Function

Private Function BuildFilePath() As String
    Dim BaseFolderPath As String
    BaseFolderPath = This.Provider.GetBaseFolderPath
    Dim FileName As String
    FileName = "UpperValue.csv"
    BuildFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "Config", FileName)
End Function
