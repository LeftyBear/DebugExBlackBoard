VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_CSVReadRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Repository")
Option Explicit

Private Type Member
    Reader  As Inf_TextStreamReader
    Parser  As Inf_CSVRFCParser
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Reader = New Inf_TextStreamReader
    Set This.Parser = New Inf_CSVRFCParser
End Sub

Public Function Load(ByVal FilePath As String) As String()
    Dim TextData As String
    TextData = This.Reader.ReadAll(FilePath)
    Load = This.Parser.Parse(TextData)
End Function
