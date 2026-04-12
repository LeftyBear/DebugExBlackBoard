VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ScheduleCSVReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Reader"
Option Explicit
Implements Dom_IScheduleReader
Private Type Member
    Repository As Inf_CSVReadRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_CSVReadRepository)
    Set This.Repository = Repository
End Sub

Private Function Dom_IScheduleReader_Load(ByVal FilePath As String) As String()
    Dim RawRows() As String
    RawRows = This.Repository.Load(FilePath)
    Dom_IScheduleReader_Load = RawRows
End Function
