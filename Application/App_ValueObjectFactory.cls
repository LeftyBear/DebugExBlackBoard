VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ValueObjectFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.Factory"
Option Explicit

Public Function CreateClassHourType(ByVal Value As Long) As Dom_ClassHourType
    With New Dom_ClassHourType
        Set CreateClassHourType = .Create(Value)
    End With
End Function

Public Function CreateStreamType(ByVal Value As Long) As Dom_StreamType
    With New Dom_StreamType
        Set CreateStreamType = .Create(Value)
    End With
End Function

Public Function CreateGrade(ByVal Value As Long) As Dom_Grade
    With New Dom_Grade
        Set CreateGrade = .Create(Value)
    End With
End Function

Public Function CreateClassNo(ByVal Value As Long) As Dom_ClassNo
    With New Dom_ClassNo
        Set CreateClassNo = .Create(Value)
    End With
End Function

Public Function CreateGender(ByVal Value As Long) As Dom_Gender
    With New Dom_Gender
        Set CreateGender = .Create(Value)
    End With
End Function

Public Function CreatePeriod(ByVal Value As Long) As Dom_Period
    With New Dom_Period
        Set CreatePeriod = .Create(Value)
    End With
End Function
