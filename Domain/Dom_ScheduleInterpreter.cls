VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Interpreter")
Option Explicit

Private Type Member
    Resolver    As Dom_HeaderTokenResolver
    Factory     As Dom_ScheduleFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
    Set This.Factory = New Dom_ScheduleFactory
End Sub

Public Function Interprete(ByVal Header As Dom_ScheduleHeader, ByVal RawText As String) As Dom_Schedule
    Dim ID As Dom_ScheduleID
    Set ID = CreateScheduleID(Header.RawText)
    Set Interprete = This.Factory.Create(ID, RawText)
End Function

Private Function CreateScheduleID(ByVal Key As String) As Dom_ScheduleID
    With New Dom_ScheduleID
        Set CreateScheduleID = .Create(Key)
    End With
End Function

