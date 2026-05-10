VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_UpperValueUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Repository  As Dom_IUpperValueRepository
    Builder     As App_UpperValueRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_UpperValueRepository, ByVal Builder As App_UpperValueRowBuilder)
    Set This.Repository = Repository
    Set This.Builder = Builder
End Sub

Public Function Create() As App_LoadUpperValueUseCase
    Dim Result As App_LoadUpperValueUseCase
    Set Result = New App_LoadUpperValueUseCase
    Result.Initialize This.Repository, New Dom_UpperValueFactory, This.Builder
    Set Create = Result
End Function

