VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ScheduleUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Repository  As Dom_IScheduleRepository
    Builder     As App_ScheduleRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_ScheduleRepository, ByVal Builder As App_ScheduleRowBuilder)
    Set This.Repository = Repository
    Set This.Builder = Builder
End Sub

Public Function Create() As App_LoadScheduleUseCase
    Dim Factory As Dom_YearlyScheduleFactory
    Set Factory = CreateYearlyScheduleFactory
    Dim Result As App_LoadScheduleUseCase
    Set Result = New App_LoadScheduleUseCase
    Result.Initialize This.Repository, Factory, This.Builder
    Set Create = Result
End Function

Private Function CreateYearlyScheduleFactory() As Dom_YearlyScheduleFactory
    Dim Factory As Dom_DailyScheduleFactory
    Set Factory = CreateDailyScheduleFactory
    Dim Result As Dom_YearlyScheduleFactory
    Set Result = New Dom_YearlyScheduleFactory
    Result.Initialize Factory
    Set CreateYearlyScheduleFactory = Result
End Function

Private Function CreateDailyScheduleFactory() As Dom_DailyScheduleFactory
    Dim Result As Dom_DailyScheduleFactory
    Set Result = New Dom_DailyScheduleFactory
    Result.Initialize New Dom_ValueObjectFactory
    Set CreateDailyScheduleFactory = Result
End Function
