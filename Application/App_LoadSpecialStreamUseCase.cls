VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadSpecialStreamUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Repository  As Dom_ISpecialStreamRepository
    Factory     As Dom_SpecialStreamFactory
    Builder     As App_SpecialStreamRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_ISpecialStreamRepository, ByVal Factory As Dom_SpecialStreamFactory, ByVal Builder As App_SpecialStreamRowBuilder)
    Set This.Repository = Repository
    Set This.Factory = Factory
    Set This.Builder = Builder
End Sub

Public Function Execute() As Dom_ISpecialStreamRepository
    Dim Records As VBA.Collection
    Set Records = This.Repository.ReadAll
    If Records.Count = 0 Then Exit Function
    Dim Rows As VBA.Collection
    Set Rows = This.Builder.Build(Records)
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Entry As Dom_SpecialStream
        Set Entry = This.Factory.Create(Rows.Item(i))
        This.Repository.Add Entry
    Next
    Set Execute = This.Repository
End Function
