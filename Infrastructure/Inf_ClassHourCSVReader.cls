VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourCSVReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Reader"
Option Explicit
Implements Dom_IClassHourReader
Private Type Member
    Repository As Inf_CSVReadRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_CSVReadRepository)
    Set This.Repository = Repository
End Sub

Private Function Dom_IClassHourReader_Load(ByVal FilePath As String) As String()
    Dim RawRows() As String
    RawRows = This.Repository.Load(FilePath)
    Dom_IClassHourReader_Load = RawRows
End Function
