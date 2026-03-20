VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourAggregateFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Private Type Member
    Parser      As Dom_ClassHourHeaderParser
    Interpreter As Dom_ClassHourInterpreter
End Type

Private This As Member

Public Sub Initialize(ByVal Parser As Dom_ClassHourHeaderParser, ByVal Interpreter As Dom_ClassHourInterpreter)
    Set This.Parser = Parser
    Set This.Interpreter = Interpreter
End Sub

Public Function Create(ByVal Config As Dom_SchoolConfig, ByVal ColumnDTOContainer As VBA.Collection) As Dom_ClassHourYearAggregate
    Dim Result As Dom_ClassHourYearAggregate
    Set Result = New Dom_ClassHourYearAggregate
    Result.Initialize Config.SubjectCount, Config.MainStreamGradeCount
    Dim Aggregate As Dom_ClassHourAggregate
    Set Aggregate = New Dom_ClassHourAggregate
    Aggregate.Initialize Dom_ClassHourType.Execution, Config.SubjectCount, Config.PeriodCount, Config.MainStreamGradeCount
    Dim ColumnDTOs As VBA.Collection
    For Each ColumnDTOs In ColumnDTOContainer
        Dim ColumnDTO As App_ColumnDTO
        For Each ColumnDTO In ColumnDTOs
            Dim Header As Dom_ClassHourHeader
            Set Header = This.Parser.Parse(Config, ColumnDTO.HeaderName)
            If Header.TargetDate <> vbNullString Then
                Dim TargetDate As Date
                TargetDate = Dom_Util_Date.NormalizeToDate(ColumnDTO.RawValue)
            Else
                This.Interpreter.Interprete Config, Header, ColumnDTO.RawValue, Aggregate
            End If
        Next
        Result.Add TargetDate, Aggregate
    Next
    Set Create = Result
End Function
