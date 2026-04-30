Attribute VB_Name = "Dom_TypePolicy"
'@Folder("Domain.Policy")
Option Explicit
Option Private Module

Public Enum Dom_ClassHourTypePolicy
    Plan = 1
    Execution
End Enum

Public Enum Dom_StreamTypePolicy
    Main = 1
    Special
End Enum

Public Enum Dom_GenderTypePolicy
    Male = 1
    Female
End Enum

Public Enum Dom_EntityTypePolicy
    ClassHour = 1
    Enrollment
    Schedule
End Enum

Public Enum Dom_StructureTypePolicy
    SchoolEvent = 1
    SpecialStream
    Subject
    UpperValues
End Enum
