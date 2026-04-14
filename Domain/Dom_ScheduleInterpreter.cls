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
    VOFactory   As Dom_ValueObjectFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
    Set This.Factory = New Dom_ScheduleFactory
    Set This.VOFactory = New Dom_ValueObjectFactory
End Sub

Public Function Interprete(ByVal Header As Dom_ScheduleHeader, ByVal RawText As String) As Dom_Schedule
    Dim ID As Dom_ScheduleID
    Set ID = This.VOFactory.CreateScheduleID(Header.RawText)
    Set Interprete = This.Factory.Create(ID, RawText)
End Function
