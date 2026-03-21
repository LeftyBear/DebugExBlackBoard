VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateClassHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Repository  As Dom_IClassHourRepository
    Config      As Dom_SchoolStructure
    Parser      As Dom_ClassHourHeaderParser
    Interpreter As Dom_ClassHourInterpreter
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IClassHourRepository, ByVal Config As Dom_SchoolStructure, ByVal Parser As Dom_ClassHourHeaderParser, ByVal Interpreter As Dom_ClassHourInterpreter)
    Set This.Repository = Repository
    Set This.Config = Config
    Set This.Parser = Parser
    Set This.Interpreter = Interpreter
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_ClassHourYearAggregate
    Dim SchoolYear As Long
    SchoolYear = Dom_Util_Date.GetSchoolYear(TargetDate)
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(SchoolYear)
    If CSVDTO.RowCount = 0 Then Err.Raise AppErrEmptyData, "App_AggregateClassHour", "CSV data is empty."
    With New App_HeaderMapper
        Dim ColumnDTOGroups As VBA.Collection
        Set ColumnDTOGroups = .Map(CSVDTO)
    End With
    Dim Factory As Dom_ClassHourAggregateFactory
    Set Factory = New Dom_ClassHourAggregateFactory
    Factory.Initialize This.Parser, This.Interpreter
    Set Execute = Factory.Create(This.Config, ColumnDTOGroups)
End Function
