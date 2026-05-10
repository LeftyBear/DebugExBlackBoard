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
    CSV As Inf_CSVPersistence
End Type

Private This As Member

Public Sub Initialize(ByVal CSV As Inf_CSVPersistence)
    Set This.CSV = CSV
End Sub

Public Function ReadAll() As VBA.Collection
    Dim FilePath As String
    FilePath = BuiledFilePath
    Set ReadAll = This.CSV.ReadAll(FilePath)
End Function

Public Sub WriteAll(ByVal SchoolYear As Long, ByVal Rows As VBA.Collection)
    Dim FilePath As String
    FilePath = BuiledFilePath
    This.CSV.WriteAll FilePath, Rows
End Sub

Public Sub Delete(ByVal SchoolYear As Long, ByVal Key As String)
    Dim FilePath As String
    FilePath = BuiledFilePath
    This.CSV.Delete FilePath, Key
End Sub

Private Function BuiledFilePath() As String
    Dim Provider As Inf_WorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim BaseFolderPath As String
    BaseFolderPath = Provider.GetBaseFolderPath
    Dim FileName As String
    FileName = "SchoolEventCatalog.csv"
    BuiledFilePath = Inf_StringUtility.JoinByBackSlash(BaseFolderPath, "root", "Config", FileName)
End Function
