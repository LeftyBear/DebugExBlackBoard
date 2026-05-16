VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    ClassHourType   As Dom_ClassHourType
    Period          As Dom_Period
    Grade           As Dom_Grade
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourType As Dom_ClassHourType, ByVal Period As Dom_Period, ByVal Grade As Dom_Grade)
    Set This.ClassHourType = ClassHourType
    Set This.Period = Period
    Set This.Grade = Grade
End Sub

Public Property Get ClassHourType() As Dom_ClassHourType
    Set ClassHourType = This.ClassHourType
End Property

Public Property Get Period() As Dom_Period
    Set Period = This.Period
End Property

Public Property Get Grade() As Dom_Grade
    Set Grade = This.Grade
End Property

Public Function ToKey() As String
    Dim ClassHourTypeMapper As Inf_ClassHourTypeMapper
    Set ClassHourTypeMapper = New Inf_ClassHourTypeMapper
    Dim PeriodMapper As Inf_PeriodMapper
    Set PeriodMapper = New Inf_PeriodMapper
    Dim GradeMapper As Inf_GradeMapper
    Set GradeMapper = New Inf_GradeMapper
    ToKey = ClassHourTypeMapper.ToRaw(This.ClassHourType) & charUnderScore & PeriodMapper.ToRaw(This.Period) & charUnderScore & GradeMapper.ToRaw(This.Grade)
End Function

Public Function Equals(ByVal Other As Dom_TimeTableID) As Boolean
    If Not This.ClassHourType.Equals(Other.ClassHourType) Then Exit Function
    If Not This.Period.Equals(Other.Period) Then Exit Function
    If Not This.Grade.Equals(Other.Grade) Then Exit Function
    Equals = True
End Function
