VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ValueObjectFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function CreateDateID(ByVal Value As Date) As Dom_DateID
    Dim VO As Dom_DateID
    Set VO = New Dom_DateID
    VO.Initialize Value
    Set CreateDateID = VO
End Function

Public Function CreateClassHourID(ByVal ClassHourType As Dom_ClassHourType, ByVal SubjectName As Dom_SubjectName, ByVal Grade As Dom_Grade) As Dom_ClassHourID
    Dim VO As Dom_ClassHourID
    Set VO = New Dom_ClassHourID
    VO.Initialize ClassHourType, SubjectName, Grade
    Set CreateClassHourID = VO
End Function

Public Function CreateTimeTableID(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade) As Dom_TimeTableID
    Dim VO As Dom_TimeTableID
    Set VO = New Dom_TimeTableID
    VO.Initialize ClassHourType, Period, Grade
    Set CreateTimeTableID = VO
End Function

Public Function CreateClassHourType(ByVal Value As Long) As Dom_ClassHourType
    Dim VO As Dom_ClassHourType
    Set VO = New Dom_ClassHourType
    VO.Initialize Value
    Set CreateClassHourType = VO
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

Public Function CreateSubjectName(ByVal Value As String) As Dom_SubjectName
    Dim VO As Dom_SubjectName
    Set VO = New Dom_SubjectName
    VO.Initialize Value
    Set CreateSubjectName = VO
End Function

Public Function CreateSubjectTargetGrade(ByVal Value As String) As Dom_SubjectTargetGrade
    Dim VO As Dom_SubjectTargetGrade
    Set VO = New Dom_SubjectTargetGrade
    VO.Initialize Value
    Set CreateSubjectTargetGrade = VO
End Function

Public Function CreateSubjectMark(ByVal Value As String) As Dom_SubjectMark
    Dim VO As Dom_SubjectMark
    Set VO = New Dom_SubjectMark
    VO.Initialize Value
    Set CreateSubjectMark = VO
End Function

Public Function CreateEnrollmentID(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender) As Dom_EnrollmentID
    Dim VO As Dom_EnrollmentID
    Set VO = New Dom_EnrollmentID
    VO.Initialize StreamType, Grade, ClassNo, Gender
    Set CreateEnrollmentID = VO
End Function

Public Function CreateStreamType(ByVal Value As Long) As Dom_StreamType
    Dim VO As Dom_StreamType
    Set VO = New Dom_StreamType
    VO.Initialize Value
    Set CreateStreamType = VO
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

Public Function CreateTransfer(ByVal Value As Boolean) As Dom_Transfer
    Dim VO As Dom_Transfer
    Set VO = New Dom_Transfer
    VO.Initialize Value
    Set CreateTransfer = VO
End Function

Public Function CreateRemarks(ByVal Value As String) As Dom_Remarks
    Dim VO As Dom_Remarks
    Set VO = New Dom_Remarks
    VO.Initialize Value
    Set CreateRemarks = VO
End Function

Public Function CreateScheduleID(ByVal Value As String) As Dom_ScheduleID
    Dim VO As Dom_ScheduleID
    Set VO = New Dom_ScheduleID
    VO.Initialize Value
    Set CreateScheduleID = VO
End Function

Public Function CreateScheduleValue(ByVal Value As String) As Dom_ScheduleValue
    Dim VO As Dom_ScheduleValue
    Set VO = New Dom_ScheduleValue
    VO.Initialize Value
    Set CreateScheduleValue = VO
End Function

Public Function CreateMainStream(ByVal UpperGrade As Long, ByVal UpperClassNo As Long, ByVal Period As Long) As Dom_MainStreamValue
    Dim VO As Dom_MainStreamValue
    Set VO = New Dom_MainStreamValue
    VO.Initialize UpperGrade, UpperClassNo
    Set CreateMainStream = VO
End Function
