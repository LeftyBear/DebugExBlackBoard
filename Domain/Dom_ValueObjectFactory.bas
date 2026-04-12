Attribute VB_Name = "Dom_ValueObjectFactory"
'@Folder "Domain.Factory"
Option Explicit
Option Private Module

Public Function CreateDateID(ByVal RawValue As String) As Dom_DateID
    Dim Value As Date
    Value = Dom_Util_Date.NormalizeToDate(RawValue)
    With New Dom_DateID
        Set CreateDateID = .Create(Value)
    End With
End Function

Public Function CreateTransfer(ByVal RawValue As String) As Dom_Transfer
    Dim Value As Boolean
    Value = Dom_Util_Boolean.NormalizeToBoolean(RawValue)
    With New Dom_Transfer
        Set CreateTransfer = .Create(Value)
    End With
End Function

Public Function CreateRemarks(ByVal RawValue As String) As Dom_Remarks
    With New Dom_Remarks
        Set CreateRemarks = .Create(RawValue)
    End With
End Function

Public Function CreateEnrollmentValue(ByVal RawValue As String) As Dom_EnrollmentValue
    Dim Value As Long
    Value = Dom_Util_Numeric.NormalizeToLong(RawValue)
    With New Dom_EnrollmentValue
        Set CreateEnrollmentValue = .Create(Value)
    End With
End Function

Public Function CreateClassHourValue(ByVal RawValue As String) As Dom_ClassHourValue
    Dim Value As Long
    Value = Dom_Util_Numeric.NormalizeToLong(RawValue)
    With New Dom_ClassHourValue
        Set CreateClassHourValue = .Create(Value)
    End With
End Function

Public Function CreateSchedlueValue(ByVal RawValue As String) As Dom_ScheduleValue
    With New Dom_ClassHourValue
        Set CreateSchedlueValue = .Create(RawValue)
    End With
End Function

Public Function CreateScheduleID(ByVal Key As String) As Dom_ScheduleID
    With New Dom_ScheduleID
        Set CreateScheduleID = .Create(Key)
    End With
End Function

Public Function CreateClassHourID(ByVal ClassHourType As Dom_ClassHourType, ByVal Subject As Dom_Subject, ByVal Grade As Dom_Grade) As Dom_ClassHourID
    With New Dom_ClassHourID
        Set CreateClassHourID = .Create(ClassHourType, Subject, Grade)
    End With
End Function

Public Function CreateEnrollmentID(ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender) As Dom_EnrollmentID
    With New Dom_EnrollmentID
        Set CreateEnrollmentID = .Create(StreamType, Grade, ClassNo, Gender)
    End With
End Function
