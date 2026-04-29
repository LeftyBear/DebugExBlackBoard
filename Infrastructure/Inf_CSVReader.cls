VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_CSVReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit

Private Type Member
    Reader  As Inf_TextStreamReader
    Parser  As Inf_CSVRFCParser
End Type

Private This As Member

Public Sub Initialize(ByVal Reader As Inf_TextStreamReader, ByVal Parser As Inf_CSVRFCParser)
    Set This.Reader = Reader
    Set This.Parser = Parser
End Sub

Public Function Load(ByVal FilePath As String) As String()
    Dim TextData As String
    TextData = This.Reader.ReadAll(FilePath)
    Load = This.Parser.Parse(TextData)
End Function
