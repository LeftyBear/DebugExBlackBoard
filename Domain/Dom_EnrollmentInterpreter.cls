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

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
End Sub

Public Function Interprete(ByVal Header As Dom_EnrollmentHeader, ByVal RawValue As String) As Dom_Enrollment
    Dim StreamType As Dom_StreamType
    Set StreamType = This.Resolver.ResolveStreamType(Header.Grade)
    Dim Grade As Dom_Grade
    Set Grade = This.Resolver.ResolveGrade(Header.Grade)
    Dim ClassNo As Dom_ClassNo
    Set ClassNo = This.Resolver.ResolveClassNo(Header.ClassNo)
    Dim Gender As Dom_Gender
    Set Gender = This.Resolver.ResolveGender(Header.Gender)
    Dim EnrollmentID As Dom_EnrollmentID
    Set EnrollmentID = CreateID(StreamType, Grade, ClassNo, Gender)
    If Header.RawDate <> vbNullString Then
        Dim NormDate As Date
        NormDate = This.Resolver.ResolveDate(RawValue)
    End If
    Dim Value As Dom_EnrollmentValue
    Set Value = CreateValue(RawValue)
    If Header.Transfer <> vbNullString Then
        Dim Transfer As Boolean
        Transfer = This.Resolver.ResolveTransfer(RawValue)
    End If
    If Header.Remarks <> vbNullString Then
        Dim Remarks As String
        Remarks = This.Resolver.ResolveRemarks(RawValue)
    End If
    Dim Enrollment As Dom_Enrollment
    Set Enrollment = New Dom_Enrollment
    Enrollment.Initialize EnrollmentID, NormDate, Value, Transfer, Remarks
    Set Interprete = Enrollment
End Function

Private Function CreateID(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender) As Dom_EnrollmentID
    With New Dom_EnrollmentID
        Set CreateID = .Create(StreamType, Grade, ClassNo, Gender)
    End With
End Function

Private Function CreateValue(ByVal Value As Long) As Dom_EnrollmentValue
    With New Dom_EnrollmentValue
        Set CreateValue = .Create(Value)
    End With
End Function
