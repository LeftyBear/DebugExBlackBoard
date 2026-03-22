VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
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
Public Sub Interprete(ByVal Config As Dom_SchoolStructure, ByVal Header As Dom_ClassHourHeader, ByVal RawValue As String, ByVal Aggregate As Dom_ClassHourAggregate)
    Dim NormDate As Date
    NormDate = This.Resolver.ResolveDate(Header.RawDate)
    Dim ClassHourType As Dom_ClassHourType
    Set ClassHourType = This.Resolver.ResolveClassHourType(Header.ClassHourType)
    If Header.SubjectName <> vbNullString Then
        Dim Subject As Dom_Subject
        Set Subject = Config.ResolveSubject(Header.SubjectName)
        Dim Hours As Dom_ClassHourValue
        Set Hours = CreateHours(Dom_Util_Long.NormalizeToLong(RawValue))
    ElseIf Header.Period <> vbNullString Then
        Dim Period As Dom_Period
        Set Period = This.Resolver.ResolvePeriod(Header.Period)
        Dim Mark As Dom_ClassHourMark
        Set Mark = CreateMark(RawValue)
    End If
    Dim Grade As Dom_Grade
    Set Grade = This.Resolver.ResolveGrade(Header.Grade)
    Dim ID As Dom_ClassHourID
    Set ID = CreateID(ClassHourType, Subject, Period, Grade)
    Dim ClassHour As Dom_ClassHour
    Set ClassHour = New Dom_ClassHour
    ClassHour.Initialize ID, NormDate, Hours, Mark
    Aggregate.AddInitialClassHour ClassHour
End Sub

Private Function CreateHours(ByVal Value As Long) As Dom_ClassHourValue
    With New Dom_ClassHourValue
        Set CreateHours = .Create(Value)
    End With
End Function

Private Function CreateMark(ByVal Value As String) As Dom_ClassHourMark
    With New Dom_ClassHourMark
        Set CreateMark = .Create(Value)
    End With
End Function

Private Function CreateID(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_ClassHourID
    With New Dom_ClassHourID
        Set CreateID = .Create(ClassHourType, Subject, Period, Grade)
    End With
End Function
