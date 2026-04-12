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
    Structure As Dom_SchoolStructure
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
    Set This.Factory = New Dom_ScheduleFactory
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Interprete(ByVal HeaderName As String, ByVal RawText As String) As Dom_Schedule
    If This.Structure.IsSchoolEvent(HeaderName) Then Exit Function
    Dim ID As Dom_ScheduleID
    Set ID = Dom_ValueObjectFactory.CreateScheduleID(HeaderName)
    Set Interprete = This.Factory.Create(ID, RawText)
End Function
