VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ConfigCSVReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit
Implements App_IConfigReader
Private Type Member
    Reader As Inf_CSVReader
End Type

Private This As Member

Public Sub Initialize(ByVal Reader As Inf_CSVReader)
    Set This.Reader = Reader
End Sub

Private Function App_IConfigReader_Load(ByVal FilePath As String) As String()
    App_IConfigReader_Load = This.Reader.Load(FilePath)
End Function
