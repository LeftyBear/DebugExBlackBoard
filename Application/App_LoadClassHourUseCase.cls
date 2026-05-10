VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadClassHourUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Repository  As Dom_IClassHourRepository
    Factory     As Dom_YearlyClassHourFactory
    Builder     As App_ClassHourRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IClassHourRepository, ByVal Factory As Dom_YearlyClassHourFactory, ByVal Builder As App_ClassHourRowBuilder)
    Set This.Repository = Repository
    Set This.Factory = Factory
    Set This.Builder = Builder
End Sub

Public Function Execute(ByVal SchoolYear As Long, ByVal Structure As Dom_SchoolStructure) As Dom_IClassHourRepository
    Dim Records As VBA.Collection
    Set Records = This.Repository.ReadAll(SchoolYear)
    If Records.Count = 0 Then Exit Function
    Dim Rows As VBA.Collection
    Set Rows = This.Builder.Build(Records, Structure)
    Dim Yearly As VBA.Collection
    Set Yearly = This.Factory.Create(Rows)
    Dim i As Long
    For i = 1 To Yearly.Count
        Dim Daily As Dom_DailyClassHour
        Set Daily = Yearly.Item(i)
        This.Repository.Add Daily
    Next
    Set Execute = This.Repository
End Function
