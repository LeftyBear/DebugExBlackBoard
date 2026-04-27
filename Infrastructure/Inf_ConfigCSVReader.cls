VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ConfigCSVReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Reader"
Option Explicit
Implements Dom_IConfigReader
Private Type Member
    Repository As Inf_CSVReader
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_CSVReader)
    Set This.Repository = Repository
End Sub

Private Function Dom_IConfigReader_Load(ByVal FilePath As String) As String()
    Dom_IConfigReader_Load = This.Repository.Load(FilePath)
End Function
