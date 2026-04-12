VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableInterpreter"
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

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
End Sub

Public Function Interprete(ByVal Header As Dom_ClassHourHeader, ByVal RawValue As String) As Dom_TimeTable
    If Header.ClassHourType <> vbNullString Then
        Dim ClassHourType As Dom_ClassHourType
        Set ClassHourType = This.Resolver.ResolveClassHourType(RawValue)
    End If
    If Header.Period <> vbNullString Then
        Dim Period As Dom_Period
        Set Period = This.Resolver.ResolvePeriod(RawValue)
    End If
    If Header.Grade <> vbNullString Then
        Dim Grade As Dom_Grade
        Set Grade = This.Resolver.ResolveGrade(RawValue)
    End If
    Dim TimeTableID As Dom_TimeTableID
    Set TimeTableID = CreateID(ClassHourType, Period, Grade)
    If TimeTableID.IsNoMeaning Then Exit Function
    Dim Value As Dom_TimeTableMark
    Set Value = CreateMark(RawValue)
    Dim TimeTable As Dom_TimeTable
    Set TimeTable = New Dom_TimeTable
    TimeTable.Initialize TimeTableID, Value
    Set Interprete = TimeTable
End Function

Private Function CreateID(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_TimeTableID
    With New Dom_TimeTableID
        Set CreateID = .Create(ClassHourType, Period, Grade)
    End With
End Function

Private Function CreateMark(ByVal RawValue As String) As Dom_TimeTableMark
    With New Dom_TimeTableMark
        Set CreateMark = .Create(RawValue)
    End With
End Function
