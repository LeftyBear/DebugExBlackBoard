VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadScheduleUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Repository  As Dom_IScheduleRepository
    Factory     As Dom_YearlyScheduleFactory
    Builder     As App_ScheduleRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IScheduleRepository, ByVal Factory As Dom_YearlyScheduleFactory, ByVal Builder As App_ScheduleRowBuilder)
    Set This.Repository = Repository
    Set This.Factory = Factory
    Set This.Builder = Builder
End Sub

Public Function Execute(ByVal SchoolYear As Long, ByVal Structure As Dom_SchoolStructure) As Dom_IScheduleRepository
    Dim Records As VBA.Collection
    Set Records = This.Repository.ReadAll(SchoolYear)
    If Records.Count = 0 Then Exit Function
    Dim Rows As VBA.Collection
    Set Rows = This.Builder.Build(Records, Structure)
    Dim Yearly As VBA.Collection
    Set Yearly = This.Factory.Create(Rows)
    Dim i As Long
    For i = 1 To Yearly.Count
        Dim Daily As Dom_DailySchedule
        Set Daily = Yearly.Item(i)
        This.Repository.Add Daily
    Next
    Set Execute = This.Repository
End Function
