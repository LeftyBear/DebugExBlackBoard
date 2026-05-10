VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Repository  As Dom_IClassHourRepository
    Builder     As App_ClassHourRowBuilder
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Inf_ClassHourRepository, ByVal Builder As App_ClassHourRowBuilder)
    Set This.Repository = Repository
    Set This.Builder = Builder
End Sub

Public Function Create() As App_LoadClassHourUseCase
    Dim Factory As Dom_YearlyClassHourFactory
    Set Factory = CreateYearlyClassHourFactory
    Dim Result As App_LoadClassHourUseCase
    Set Result = New App_LoadClassHourUseCase
    Result.Initialize This.Repository, Factory, This.Builder
    Set Create = Result
End Function

Private Function CreateYearlyClassHourFactory() As Dom_YearlyClassHourFactory
    Dim ClassHourFactory As Dom_DailyClassHourFactory
    Set ClassHourFactory = CreateDailyClassHourFactory
    Dim TimeTableFactory As Dom_DailyTimeTableFactory
    Set TimeTableFactory = CreateDailyTimeTableFactory
    Dim Result As Dom_YearlyClassHourFactory
    Set Result = New Dom_YearlyClassHourFactory
    Result.Initialize ClassHourFactory, TimeTableFactory
    Set CreateYearlyClassHourFactory = Result
End Function

Private Function CreateDailyClassHourFactory() As Dom_DailyClassHourFactory
    Dim Result As Dom_DailyClassHourFactory
    Set Result = New Dom_DailyClassHourFactory
    Result.Initialize New Dom_ValueObjectFactory
    Set CreateDailyClassHourFactory = Result
End Function

Private Function CreateDailyTimeTableFactory() As Dom_DailyTimeTableFactory
    Dim Result As Dom_DailyTimeTableFactory
    Set Result = New Dom_DailyTimeTableFactory
    Result.Initialize New Dom_ValueObjectFactory
    Set CreateDailyTimeTableFactory = Result
End Function
