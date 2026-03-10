VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Repository"
Option Explicit
Implements Dom_ISubjectRepository
Private Type Member
    Repository As Inf_ConfigReadRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_ConfigReadRepository)
    Set This.Repository = Repository
End Sub

Private Function Dom_ISubjectRepository_Load(ByVal ConfigName As Dom_Config) As String()
    Dom_ISubjectRepository_Load = This.Repository.Load(ConfigName)
End Function
