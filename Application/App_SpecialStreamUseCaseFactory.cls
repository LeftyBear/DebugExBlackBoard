VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SpecialStreamUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Repository  As Dom_ISpecialStreamRepository
    Builder     As App_SpecialStreamRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_SpecialStreamRepository, ByVal Builder As App_SpecialStreamRowBuilder)
    Set This.Repository = Repository
    Set This.Builder = Builder
End Sub

Public Function Create() As App_LoadSpecialStreamUseCase
    Dim Result As App_LoadSpecialStreamUseCase
    Set Result = New App_LoadSpecialStreamUseCase
    Result.Initialize This.Repository, New Dom_SpecialStreamFactory, This.Builder
    Set Create = Result
End Function
