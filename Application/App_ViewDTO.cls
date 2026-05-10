VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ViewDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.DTO")
Option Explicit

Private Type Member
    Structure                   As Dom_SchoolStructure
    Schedule                    As Dom_DailySchedule
    EnrollmentTable()           As Variant
    ClassHourPlanTable()        As Variant
    ClassHourExecutionTable()   As Variant
    TimeTablePlanTable()        As Variant
    TimeTableExecutionTable()   As Variant
End Type

Private This As Member

Public Property Set Structure(ByVal Value As Dom_SchoolStructure)
    Set This.Structure = Value
End Property

Public Property Set Schedule(ByVal Value As Dom_DailySchedule)
    Set This.Schedule = Value
End Property

Public Property Let EnrollmentTable(ByRef Value() As Variant)
    This.EnrollmentTable = Value
End Property

Public Property Let ClassHourPlanTable(ByRef Value() As Variant)
    This.ClassHourPlanTable = Value
End Property

Public Property Let ClassHourExecutionTable(ByRef Value() As Variant)
    This.ClassHourExecutionTable = Value
End Property

Public Property Let TimeTablePlanTable(ByRef Value() As Variant)
    This.TimeTablePlanTable = Value
End Property

Public Property Let TimeTableExecutionTable(ByRef Value() As Variant)
    This.TimeTableExecutionTable = Value
End Property

Public Property Get Structure() As Dom_SchoolStructure
    Set Structure = This.Structure
End Property

Public Property Get Schedule() As Dom_DailySchedule
    Set Schedule = This.Schedule
End Property

Public Property Get EnrollmentTable() As Variant()
    EnrollmentTable = This.EnrollmentTable
End Property

Public Property Get ClassHourPlanTable() As Variant()
    ClassHourPlanTable = This.ClassHourPlanTable
End Property

Public Property Get ClassHourExecutionTable() As Variant()
    ClassHourExecutionTable = This.ClassHourExecutionTable
End Property

Public Property Get TimeTablePlanTable() As Variant()
    TimeTablePlanTable = This.TimeTablePlanTable
End Property

Public Property Get TimeTableExecutionTable() As Variant()
    TimeTableExecutionTable = This.TimeTableExecutionTable
End Property
