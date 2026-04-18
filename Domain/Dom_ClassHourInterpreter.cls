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
    Resolver    As Dom_HeaderTokenResolver
    Factory     As Dom_ClassHourFactory
    Structure   As Dom_SchoolStructure
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
    Set This.Factory = New Dom_ClassHourFactory
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Interprete(ByVal Header As Dom_ClassHourHeader, ByVal RawValue As String) As Dom_ClassHour
    If Header.ClassHourType <> vbNullString Then
        Dim ClassHourType As Dom_ClassHourType
        Set ClassHourType = This.Resolver.ResolveClassHourType(RawValue)
    End If
    If Header.SubjectName <> vbNullString Then
        Dim Subject As Dom_Subject
        Set Subject = This.Structure.ResolveSubject(RawValue)
    End If
    If Header.Grade <> vbNullString Then
        Dim Grade As Dom_Grade
        Set Grade = This.Resolver.ResolveGrade(RawValue)
    End If
    Dim ID As Dom_ClassHourID
    Set ID = CreateClassHourID(ClassHourType, Subject, Grade)
    If ID.IsNoMeaning Then Exit Function
    Set Interprete = This.Factory.Create(ID, RawValue)
End Function

Private Function CreateClassHourID(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Dom_ClassHourID
    With New Dom_ClassHourID
        Set CreateClassHourID = .Create(ClassHourType, Subject, Grade)
    End With
End Function

