VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Interpreter")
Option Explicit

Private Type Member
    Resolver As Dom_HeaderTokenResolver
End Type

Private This As Member

Public Sub Initialize(ByVal Resolver As Dom_HeaderTokenResolver)
    Set This.Resolver = Resolver
End Sub
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Interprete(ByVal Config As Dom_SchoolConfig, ByVal Header As Dom_ClassHourHeader, ByVal RawValue As String, ByVal Aggregate As Dom_ClassHourAggregate)
    Dim ClassHourType As Dom_ClassHourType
    ClassHourType = This.Resolver.ResolveClassHourType(Header.ClassHourType)
    Dim Grade As Dom_Grade
    Set Grade = This.Resolver.ResolveGrade(Header.Grade)
    If Header.SubjectName <> vbNullString Then
        Dim Subject As Dom_Subject
        Set Subject = Config.ResolveSubject(Header.SubjectName)
        Aggregate.AddHours ClassHourType, Subject, Grade, Dom_Util_Long.NormalizeToLong(RawValue)
    ElseIf Header.Period <> vbNullString Then
        Dim Period As Dom_Period
        Set Period = This.Resolver.ResolvePeriod(Header.Period)
        Aggregate.AddMark ClassHourType, Period, Grade, RawValue
    End If
End Sub
