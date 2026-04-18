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
    Factory         As Dom_DateIDFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.HeaderResolver = New Dom_HeaderTokenResolver
    Set This.Interpreter = New Dom_EnrollmentInterpreter
    Set This.Factory = New Dom_DateIDFactory
End Sub

'@Ignore ParameterCanBeByVal
'@Ignore ParameterCanBeByVal
Public Sub Resolve(ByVal Header As Dom_EnrollmentHeader, ByVal RawText As String, ByRef DateID As Dom_DateID, ByRef Transfer As Dom_Transfer, ByRef Remarks As Dom_Remarks, ByRef MainStreamList As Scripting.Dictionary, ByRef SpecialStreamList As Scripting.Dictionary)
    If Header.RawDate <> vbNullString Then
        Dim NormDate As Date
        NormDate = Dom_Util_Date.NormalizeToDate(RawText)
        Set DateID = This.Factory.Create(NormDate)
    ElseIf Header.Transfer <> vbNullString Then
        Set Transfer = CreateTransfer(RawText)
    ElseIf Header.Remarks <> vbNullString Then
        Set Remarks = CreateRemarks(RawText)
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

Private Function CreateTransfer(ByVal RawText As String) As Dom_Transfer
    Dim Value As Boolean
    Value = Dom_Util_Boolean.NormalizeToBoolean(RawText)
    With New Dom_Transfer
        Set CreateTransfer = .Create(Value)
    End With
End Function

Private Function CreateRemarks(ByVal RawText As String) As Dom_Remarks
    With New Dom_Remarks
        Set CreateRemarks = .Create(RawText)
    End With
End Function
