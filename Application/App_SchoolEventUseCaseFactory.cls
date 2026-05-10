VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SchoolEventUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Repository  As Dom_ISchoolEventRepository
    Builder     As App_SchoolEventRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_SchoolEventRepository, ByVal Builder As App_SchoolEventRowBuilder)
    Set This.Repository = Repository
    Set This.Builder = Builder
End Sub

Public Function Create() As App_LoadSchoolEventUseCase
    Dim Result As App_LoadSchoolEventUseCase
    Set Result = New App_LoadSchoolEventUseCase
    Result.Initialize This.Repository, New Dom_SchoolEventFactory, This.Builder
    Set Create = Result
End Function
