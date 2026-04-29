Attribute VB_Name = "Dom_TypeEnum"
'@Folder("Domain.ValueObject")
Option Explicit
Option Private Module

Public Enum Dom_ClassHourTypeEnum
    Plan = 1
    Execution
End Enum

Public Enum Dom_StreamTypeEnum
    Main = 1
    Special
End Enum

Public Enum Dom_GenderTypeEnum
    Male = 1
    Female
End Enum

Public Enum Dom_EntityTypeEnum
    ClassHour = 1
    Enrollment
    Schedule
End Enum

Public Enum Dom_StructureTypeEnum
    SchoolEvent = 1
    SpecialStream
    Subject
    UpperValues
End Enum
