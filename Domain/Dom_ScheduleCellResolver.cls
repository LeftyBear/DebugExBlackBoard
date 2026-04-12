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
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Interpreter = New Dom_ScheduleInterpreter
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    This.Interpreter.AddOn Structure
End Sub

'@Ignore ParameterCanBeByVal
'@Ignore ParameterCanBeByVal
Public Sub Resolve(ByVal HeaderName As String, ByVal RawText As String, ByRef DateID As Dom_DateID, ByRef ItemList As Scripting.Dictionary)
    If HeaderName = jpnHizuke Then
        Set DateID = Dom_ValueObjectFactory.CreateDateID(RawText)
    Else
        Dim Schedule As Dom_Schedule
        Set Schedule = This.Interpreter.Interprete(HeaderName, RawText)
        ItemList.Add HeaderName, Schedule
    End If
End Sub
