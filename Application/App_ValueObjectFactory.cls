VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ValueObjectFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Public Function CreateSchoolYear(ByVal Value As Long) As Dom_SchoolYear
    Dim VO As Dom_SchoolYear
    Set VO = New Dom_SchoolYear
    VO.Initialize Value
    Set CreateSchoolYear = VO
End Function

Public Function CreateDateID(ByVal Value As Date) As Dom_DateID
    Dim VO As Dom_DateID
    Set VO = New Dom_DateID
    VO.Initialize Value
    Set CreateDateID = VO
End Function

Public Function CreateClassHourID(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Dom_ClassHourID
    Dim VO As Dom_ClassHourID
    Set VO = New Dom_ClassHourID
    VO.Initialize ClassHourType, Subject, Grade
    Set CreateClassHourID = VO
End Function

Public Function CreateClassHourFilterCondition(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As App_ClassHourFilterCondition
    Dim VO As App_ClassHourFilterCondition
    Set VO = New App_ClassHourFilterCondition
    VO.Initialize ClassHourType, Subject, Grade
    Set CreateClassHourFilterCondition = VO
End Function

Public Function CreateTimeTableID(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_TimeTableID
    Dim VO As Dom_TimeTableID
    Set VO = New Dom_TimeTableID
    VO.Initialize ClassHourType, Period, Grade
    Set CreateTimeTableID = VO
End Function

Public Function CreateTimeTableFilterCondition(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As App_TimeTableFilterCondition
    Dim VO As App_TimeTableFilterCondition
    Set VO = New App_TimeTableFilterCondition
    VO.Initialize ClassHourType, Period, Grade
    Set CreateTimeTableFilterCondition = VO
End Function

Public Function CreateClassHourType(ByVal Value As Long) As Dom_ClassHourType
    Dim VO As Dom_ClassHourType
    Set VO = New Dom_ClassHourType
    VO.Initialize Value
    Set CreateClassHourType = VO
End Function

Public Function CreateSubject(ByVal Index As Long, ByVal Name As String, ByVal TargetGrade As String, ByVal Mark As String) As Dom_Subject
    Dim VO As Dom_Subject
    Set VO = New Dom_Subject
    VO.Initialize Index, Name, TargetGrade, Mark
    Set CreateSubject = VO
End Function

Public Function CreatePeriod(ByVal Value As Long) As Dom_Period
    Dim VO As Dom_Period
    Set VO = New Dom_Period
    VO.Initialize Value
    Set CreatePeriod = VO
End Function

Public Function CreateClassHourValue(ByVal Value As Long) As Dom_ClassHourValue
    Dim VO As Dom_ClassHourValue
    Set VO = New Dom_ClassHourValue
    VO.Initialize Value
    Set CreateClassHourValue = VO
End Function

Public Function CreateTimeTableValue(ByVal Value As Long) As Dom_TimeTableValue
    Dim VO As Dom_TimeTableValue
    Set VO = New Dom_TimeTableValue
    VO.Initialize Value
    Set CreateTimeTableValue = VO
End Function

Public Function CreateEnrollmentID(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender) As Dom_EnrollmentID
    Dim VO As Dom_EnrollmentID
    Set VO = New Dom_EnrollmentID
    VO.Initialize StreamType, Grade, ClassNo, Gender
    Set CreateEnrollmentID = VO
End Function

Public Function CreateEnrollmentFilterCondition(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, Gender As Dom_Gender) As App_EnrollmentFilterCondition
    Dim VO As App_EnrollmentFilterCondition
    Set VO = New App_EnrollmentFilterCondition
    VO.Initialize StreamType, Grade, ClassNo, Gender
    Set CreateEnrollmentFilterCondition = VO
End Function

Public Function CreateStreamType(ByVal Value As Long) As Dom_StreamType
    Dim VO As Dom_StreamType
    Set VO = New Dom_StreamType
    VO.Initialize Value
    Set CreateStreamType = VO
End Function

Public Function CreateSpecialStream(ByVal Index As Long, ByVal Name As String) As Dom_SpecialStream
    Dim VO As Dom_SpecialStream
    Set VO = New Dom_SpecialStream
    VO.Initialize Index, Name
    Set CreateSpecialStream = VO
End Function

Public Function CreateEnrollmentValue(ByVal Value As Long) As Dom_EnrollmentValue
    Dim VO As Dom_EnrollmentValue
    Set VO = New Dom_EnrollmentValue
    VO.Initialize Value
    Set CreateEnrollmentValue = VO
End Function

Public Function CreateGrade(ByVal Value As Long) As Dom_Grade
    Dim VO As Dom_Grade
    Set VO = New Dom_Grade
    VO.Initialize Value
    Set CreateGrade = VO
End Function

Public Function CreateClassNo(ByVal Value As Long) As Dom_ClassNo
    Dim VO As Dom_ClassNo
    Set VO = New Dom_ClassNo
    VO.Initialize Value
    Set CreateClassNo = VO
End Function

Public Function CreateGender(ByVal Value As Long) As Dom_Gender
    Dim VO As Dom_Gender
    Set VO = New Dom_Gender
    VO.Initialize Value
    Set CreateGender = VO
End Function

Public Function CreateSchedlueValue(ByVal Value As String) As Dom_ScheduleValue
    Dim VO As Dom_ScheduleValue
    Set VO = New Dom_ScheduleValue
    VO.Initialize Value
    Set CreateSchedlueValue = VO
End Function

Public Function CreateSchoolEvent(ByVal Index As Long, ByVal Name As String) As Dom_SchoolEvent
    Dim VO As Dom_SchoolEvent
    Set VO = New Dom_SchoolEvent
    VO.Initialize Index, Name
    Set CreateSchoolEvent = VO
End Function

Public Function CreateUpperValue(ByVal MainStreamGradeCount As Long, ByVal MainStreamClassCount As Long, ByVal PeriodCount As Long) As Dom_UpperValue
    Dim VO As Dom_UpperValue
    Set VO = New Dom_UpperValue
    VO.Initialize MainStreamGradeCount, MainStreamClassCount, PeriodCount
    Set CreateUpperValue = VO
End Function
