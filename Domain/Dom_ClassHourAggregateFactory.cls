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

Public Function Create(ByVal Config As Dom_SchoolStructure, ByVal ColumnDTOGroups As VBA.Collection) As Dom_ClassHourYearAggregate
    Dim Result As Dom_ClassHourYearAggregate
    Set Result = New Dom_ClassHourYearAggregate
'    Result.Initialize Config.SubjectCount, Config.MainStreamGradeCount
    Dim Daily As Dom_ClassHourAggregate
    Set Daily = New Dom_ClassHourAggregate
    Dim ColumnDTOGroup As VBA.Collection
    For Each ColumnDTOGroup In ColumnDTOGroups
        Dim ColumnDTO As App_ColumnDTO
        For Each ColumnDTO In ColumnDTOGroup
            Dim Header As Dom_ClassHourHeader
            Set Header = This.Parser.Parse(Config, ColumnDTO.HeaderName)
            This.Interpreter.Interprete Config, Header, ColumnDTO.RawValue, Daily
        Next
        Result.Add Daily
    Next
    Set Create = Result
End Function
