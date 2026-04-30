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
Implements App_ICSVReader

Private Function App_ICSVReader_Load(ByVal FilePath As String) As String()
    Dim Reader As Inf_TextStreamReader
    Set Reader = New Inf_TextStreamReader
    Dim TextData As String
    TextData = Reader.ReadAll(FilePath)
    Dim Parser As Inf_CSVRFCParser
    Set Parser = New Inf_CSVRFCParser
    App_ICSVReader_Load = Parser.Parse(TextData)
End Function
