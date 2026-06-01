Attribute VB_Name = "Dom_TimeTableIDFactory"
'@Folder("Domain.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_TimeTableID
    Dim ID As Dom_TimeTableID
    Set ID = New Dom_TimeTableID
    ID.Initialize ClassHourType, Period, Grade
    Set Create = ID
End Function
