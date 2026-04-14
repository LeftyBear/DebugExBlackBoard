VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentCellResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    HeaderResolver  As Dom_HeaderTokenResolver
    Interpreter     As Dom_EnrollmentInterpreter
    VOFactory       As Dom_ValueObjectFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.HeaderResolver = New Dom_HeaderTokenResolver
    Set This.Interpreter = New Dom_EnrollmentInterpreter
    Set This.VOFactory = New Dom_ValueObjectFactory
End Sub

'@Ignore ParameterCanBeByVal
'@Ignore ParameterCanBeByVal
Public Sub Resolve(ByVal Header As Dom_EnrollmentHeader, ByVal RawText As String, ByRef DateID As Dom_DateID, ByRef Transfer As Dom_Transfer, ByRef Remarks As Dom_Remarks, ByRef MainStreamList As Scripting.Dictionary, ByRef SpecialStreamList As Scripting.Dictionary)
    If Header.RawDate <> vbNullString Then
        Set DateID = This.VOFactory.CreateDateID(RawText)
    ElseIf Header.Transfer <> vbNullString Then
        Set Transfer = This.VOFactory.CreateTransfer(RawText)
    ElseIf Header.Remarks <> vbNullString Then
        Set Remarks = This.VOFactory.CreateRemarks(RawText)
    Else
        Dim Enrollment As Dom_Enrollment
        Set Enrollment = This.Interpreter.Interprete(Header, RawText)
        If Header.Grade <> jpnTokushi Then
            MainStreamList.Add ToKey(Header), Enrollment
        Else
            SpecialStreamList.Add ToKey(Header), Enrollment
        End If
    End If
End Sub

Private Function ToKey(ByVal Header As Dom_EnrollmentHeader) As String
    With Header
        ToKey = .Grade & charUnderScore & .ClassNo & charUnderScore & .Gender
    End With
End Function
