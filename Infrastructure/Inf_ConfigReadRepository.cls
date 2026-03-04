VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ConfigReadRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit

Private Type Member
    Builder As Inf_ConfigFilePathBuilder
    Reader  As Inf_TextStreamReader
    Parser  As Inf_CSVRFCParser
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As Inf_ConfigFilePathBuilder, ByVal Reader As Inf_TextStreamReader, ByVal Parser As Inf_CSVRFCParser)
    Set This.Builder = Builder
    Set This.Reader = Reader
    Set This.Parser = Parser
End Sub

Public Function Load(ByVal ConfigName As Dom_Config) As String()
    Dim FilePath As String
    FilePath = This.Builder.Builed(ConfigName)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise InfErrNotFoundFile, "", "Not found file"
    Dim TextData As String
    TextData = This.Reader.ReadAll(FilePath)
    Load = This.Parser.Parse(TextData)
End Function

