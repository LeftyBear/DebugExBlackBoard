VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EditerUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Base        As App_BaseUseCase
    RepoDTO     As App_RepositoryDTO
End Type

Private This As Member

Friend Sub Inject(ByVal Base As App_BaseUseCase, ByVal RepoDTO As App_RepositoryDTO)
    Set This.Base = Base
    Set This.RepoDTO = RepoDTO
End Sub
