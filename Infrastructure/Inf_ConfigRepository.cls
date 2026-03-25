VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ConfigRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Repository"
Option Explicit
Implements Dom_IConfigRepository
Private Type Member
    Repository As Inf_ConfigReadRepository
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Repository = New Inf_ConfigReadRepository
End Sub

Private Function Dom_IConfigRepository_Load(ByVal FilePath As String) As String()
    Dom_IConfigRepository_Load = This.Repository.Load(FilePath)
End Function
