VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregateFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Private Type Member
    Parser      As Dom_EnrollmentHeaderParser
    Interpreter As Dom_EnrollmentInterpreter
End Type

Private This As Member

Public Sub Initialize(ByVal Parser As Dom_EnrollmentHeaderParser, ByVal Interpreter As Dom_EnrollmentInterpreter)
    Set This.Parser = Parser
    Set This.Interpreter = Interpreter
End Sub

Public Function Create(ByVal Config As Dom_SchoolConfig, ByVal ColumnDTOContainer As VBA.Collection) As Dom_EnrollmentYearAggregate
    Dim GradeCount As Long
    GradeCount = Bigger(Config.MainStreamGradeCount, Config.SpecialStreamGradeCount)
    Dim ClassNoCount As Long
    ClassNoCount = Bigger(Config.MainStreamClassCount, Config.SpecialStreamClassCount)
    Dim Result As Dom_EnrollmentYearAggregate
    Set Result = New Dom_EnrollmentYearAggregate
    Result.Initialize Dom_StreamType.Special, GradeCount, ClassNoCount, Dom_Gender.Female
    Dim Aggregate As Dom_EnrollmentAggregate
    Set Aggregate = New Dom_EnrollmentAggregate
    Aggregate.Initialize Dom_StreamType.Special, GradeCount, ClassNoCount, Dom_Gender.Female
    Dim ColumnDTOs As VBA.Collection
    For Each ColumnDTOs In ColumnDTOContainer
        Dim ColumnDTO As App_ColumnDTO
        For Each ColumnDTO In ColumnDTOs
            Dim Header As Dom_EnrollmentHeader
            Set Header = This.Parser.Parse(Config, ColumnDTO.HeaderName)
            If Header.TargetDate <> vbNullString Then
                Dim TargetDate As Date
                TargetDate = Dom_Util_Date.NormalizeToDate(ColumnDTO.RawValue)
            ElseIf Header.Transfer <> vbNullString Then
                Dim Transfer As Boolean
                Transfer = Dom_Util_Boolean.NormalizeToBoolean(ColumnDTO.RawValue)
            ElseIf Header.Remarks <> vbNullString Then
                Dim Remarks As String
                Remarks = ColumnDTO.RawValue
            Else
                This.Interpreter.Interprete Config, Header, ColumnDTO.RawValue, Aggregate
            End If
        Next
        Result.Add TargetDate, Aggregate
    Next
    Set Create = Result
End Function

Private Function Bigger(ByVal Value As Long, ByVal Another As Long) As Long
    If Value < Another Then
        Bigger = Another
    Else
        Bigger = Value
    End If
End Function
