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
    Structure   As Dom_SchoolStructure
    Resolver    As Dom_HeaderTokenResolver
End Type

Private This As Member

Public Sub Initialize(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
    Set This.Resolver = New Dom_HeaderTokenResolver
End Sub

Public Function Interprete(ByVal Header As Dom_ClassHourHeader, ByVal RawValue As String) As Dom_ClassHour
    If Header.ClassHourType <> vbNullString Then
        Dim ClassHourType As Dom_ClassHourType
        Set ClassHourType = This.Resolver.ResolveClassHourType(RawValue)
    End If
    If Header.SubjectName <> vbNullString Then
        Dim Subject As Dom_Subject
        Set Subject = This.Structure.ResolveSubject(RawValue)
        Dim Hours As Dom_ClassHourValue
        Set Hours = CreateHours(Dom_Util_Long.NormalizeToLong(RawValue))
    End If
    If Header.Period <> vbNullString Then
        Dim Period As Dom_Period
        Set Period = This.Resolver.ResolvePeriod(RawValue)
        Dim Mark As Dom_ClassHourMark
        Set Mark = CreateMark(RawValue)
    End If
    If Header.Grade <> vbNullString Then
        Dim Grade As Dom_Grade
        Set Grade = This.Resolver.ResolveGrade(RawValue)
    End If
    Dim ClassHourID As Dom_ClassHourID
    Set ClassHourID = CreateID(ClassHourType, Subject, Period, Grade)
    If Header.RawDate <> vbNullString Then
        Dim NormDate As Date
        NormDate = This.Resolver.ResolveDate(RawValue)
    End If
    Dim ClassHour As Dom_ClassHour
    Set ClassHour = New Dom_ClassHour
    ClassHour.Initialize ClassHourID, NormDate, Hours, Mark
    Set Interprete = ClassHour
End Function

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
