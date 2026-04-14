VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Interpreter")
Option Explicit

Private Type Member
    Resolver    As Dom_HeaderTokenResolver
    Factory     As Dom_EnrollmentFactory
    VOFactory   As Dom_ValueObjectFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
    Set This.Factory = New Dom_EnrollmentFactory
    Set This.VOFactory = New Dom_ValueObjectFactory
End Sub

Public Function Interprete(ByVal Header As Dom_EnrollmentHeader, ByVal RawText As String) As Dom_Enrollment
    Dim StreamType As Dom_StreamType
    Set StreamType = This.Resolver.ResolveStreamType(Header.Grade)
    Dim Grade As Dom_Grade
    Set Grade = This.Resolver.ResolveGrade(Header.Grade)
    Dim ClassNo As Dom_ClassNo
    Set ClassNo = This.Resolver.ResolveClassNo(Header.ClassNo)
    Dim Gender As Dom_Gender
    Set Gender = This.Resolver.ResolveGender(Header.Gender)
    Dim ID As Dom_EnrollmentID
    Set ID = This.VOFactory.CreateEnrollmentID(StreamType, Grade, ClassNo, Gender)
    If ID.IsNoMeaning Then Exit Function
    Set Interprete = This.Factory.Create(ID, RawText)
End Function
