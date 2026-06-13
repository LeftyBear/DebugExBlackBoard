VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectPersistence"
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

Public Function ReadAll() As Inf_SubjectRows
    Dim FilePath As String
    FilePath = BuildFilePath
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Exit Function
    Dim RawRows As Inf_RawRows
    Set RawRows = This.CSV.ReadAll(FilePath)
    Set ReadAll = Inf_SubjectRowsFactory.Create(RawRows)
End Function

Public Sub WriteAll(ByVal Rows As Inf_SubjectRows)
    Dim Header() As Variant
    Header = GetHeaders
    Dim Body As VBA.Collection
    Set Body = GroupByRecord(Rows)
    Dim RawRows As Inf_RawRows
    Set RawRows = BuildRows(Header, Body)
    Dim FilePath As String
    FilePath = BuildFilePath
    This.CSV.WriteAll FilePath, RawRows
End Sub

Public Sub Delete(ByVal Key As String)
    Dim FilePath As String
    FilePath = BuildFilePath
    This.CSV.Delete FilePath, Key
End Sub

Public Function GetHeaders() As Variant()
    Dim FilePath As String
    FilePath = BuildFilePath
    GetHeaders = This.CSV.ReadAll(FilePath).GetHeaders
End Function

Private Function GroupByRecord(ByVal Rows As Inf_SubjectRows) As VBA.Collection
    Dim Records As VBA.Collection
    Set Records = New VBA.Collection
    Dim Fields As VBA.Collection
    Set Fields = New VBA.Collection
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SubjectRow
        Set Row = Rows.Item(i)
        If Row.Name <> vbNullString Then
            Fields.Add Row.Name
        ElseIf Row.TargetGrade <> vbNullString Then
            Fields.Add Row.TargetGrade
        ElseIf Row.Mark <> vbNullString Then
            Fields.Add Row.Mark
        End If
        If Fields.Count = 3 Then
            Records.Add Fields
            Set Fields = New VBA.Collection
        End If
    Next
    Set GroupByRecord = Records
End Function

Private Function BuildRows(ByRef Header() As Variant, ByVal Body As VBA.Collection) As Inf_RawRows
    Dim Result As Inf_RawRows
    Set Result = New Inf_RawRows
    Dim i As Long
    For i = LBound(Header) To UBound(Header)
        Result.AddRow Header(i)
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
    FileName = "SubjectCatalog.csv"
    BuildFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "Config", FileName)
End Function
