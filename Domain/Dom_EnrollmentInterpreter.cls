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
    Resolver As Dom_HeaderTokenResolver
End Type

Private This As Member

Public Sub Initialize(ByVal Resolver As Dom_HeaderTokenResolver)
    Set This.Resolver = Resolver
End Sub
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Interprete(ByVal Config As Dom_SchoolConfig, ByVal Header As Dom_EnrollmentHeader, ByVal RawValue As String, ByVal Aggregate As Dom_EnrollmentAggregate)
    Dim StreamType As Dom_StreamType
    StreamType = This.Resolver.ResolveStreamType(Header.Grade)
    Dim Grade As Dom_Grade
    Set Grade = This.Resolver.ResolveGrade(Header.Grade)
    Dim ClassNo As Dom_ClassNo
    Set ClassNo = This.Resolver.ResolveClassNo(Header.ClassNo)
    Dim Gender As Dom_Gender
    Gender = This.Resolver.ResolveGender(Header.Gender)
    Aggregate.AddValue StreamType, Grade, ClassNo, Gender, Dom_Util_Long.NormalizeToLong(RawValue)
End Sub
