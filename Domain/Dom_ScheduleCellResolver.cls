VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleCellResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    Interpreter As Dom_ScheduleInterpreter
    Factory     As Dom_DateIDFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Interpreter = New Dom_ScheduleInterpreter
    Set This.Factory = New Dom_DateIDFactory
End Sub

'@Ignore ParameterCanBeByVal
'@Ignore ParameterCanBeByVal
Public Sub Resolve(ByVal Header As Dom_ScheduleHeader, ByVal RawText As String, ByRef DateID As Dom_DateID, ByRef ItemList As Scripting.Dictionary)
    If Header.RawDate <> vbNullString Then
        Dim NormDate As Date
        NormDate = Dom_Util_Date.NormalizeToDate(RawText)
        Set DateID = This.Factory.Create(NormDate)
    ElseIf Header.RawText <> vbNullString Then
        Dim Schedule As Dom_Schedule
        Set Schedule = This.Interpreter.Interprete(Header, RawText)
        ItemList.Add Header.RawText, Schedule
    End If
End Sub
