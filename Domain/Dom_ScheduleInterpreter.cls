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
    Resolver As Dom_HeaderTokenResolver
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
End Sub

Public Function Interprete(ByVal HeaderName As String, ByVal RawValue As String) As Dom_Schedule
    Dim ScheduleID As Dom_ScheduleID
    Set ScheduleID = CreateID(HeaderName)
    Dim Schedule As Dom_Schedule
    Set Schedule = New Dom_Schedule
    Schedule.Initialize ScheduleID, RawValue
    Set Interprete = Schedule
End Function

Private Function CreateID(ByVal Key As String) As Dom_ScheduleID
    With New Dom_ScheduleID
        Set CreateID = .Create(Key)
    End With
End Function
