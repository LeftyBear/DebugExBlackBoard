VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SubjectUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Repository  As Dom_ISubjectRepository
    Builder     As App_SubjectRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_SubjectRepository, ByVal Builder As App_SubjectRowBuilder)
    Set This.Repository = Repository
    Set This.Builder = Builder
End Sub

Public Function Create() As App_LoadSubjectUseCase
    Dim Result As App_LoadSubjectUseCase
    Set Result = New App_LoadSubjectUseCase
    Result.Initialize This.Repository, New Dom_SubjectFactory, This.Builder
    Set Create = Result
End Function

