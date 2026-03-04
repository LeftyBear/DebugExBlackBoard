VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_LimitValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    MainStreamGradeCount    As Long
    MainStreamClassCount    As Long
    SpecialGradeCount       As Long
    SpecialClassCount       As Long
    PeriodCount             As Long
End Type

Private This As Member

Public Sub Initialize(ByVal MainStreamGradeCount As Long, ByVal MainStreamClassCount As Long, ByVal SpecialGradeCount As Long, ByVal SpecialClassCount As Long, ByVal PeriodCount As Long)
    This.MainStreamGradeCount = MainStreamGradeCount
    This.MainStreamClassCount = MainStreamClassCount
    This.SpecialGradeCount = SpecialGradeCount
    This.SpecialClassCount = SpecialClassCount
    This.PeriodCount = PeriodCount
End Sub

Public Property Get MainStreamGradeCount() As Long
    MainStreamGradeCount = This.MainStreamGradeCount
End Property

Public Property Get MainStreamClassCount() As Long
    MainStreamClassCount = This.MainStreamClassCount
End Property

Public Property Get SpecialGradeCount() As Long
    SpecialGradeCount = This.SpecialGradeCount
End Property

Public Property Get SpecialClassCount() As Long
    SpecialClassCount = This.SpecialClassCount
End Property

Public Property Get PeriodCount() As Long
    PeriodCount = This.PeriodCount
End Property

Public Function ToMainStreamGradeIndex(ByVal Grade As Long) As Long
    If Not MainStreamGradeContains(Grade) Then Err.Raise AppErrInvalidRange, , "Grade out of range."
    ToMainStreamGradeIndex = Grade
End Function

Public Function ToMainStreamClassIndex(ByVal ClassNo As Long) As Long
    If Not MainStreamClassContains(ClassNo) Then Err.Raise AppErrInvalidRange, , "Class out of range."
    ToMainStreamClassIndex = ClassNo
End Function

Public Function ToSpecialGradeIndex(ByVal Grade As Long) As Long
    If Not SpecialGradeContains(Grade) Then Err.Raise AppErrInvalidRange, , "Grade out of range."
    ToSpecialGradeIndex = Grade
End Function

Public Function ToSpecialClassIndex(ByVal ClassNo As Long) As Long
    If Not SpecialClassContains(ClassNo) Then Err.Raise AppErrInvalidRange, , "Class out of range."
    ToSpecialClassIndex = ClassNo
End Function

Public Function ToPeriodIndex(ByVal Period As Long) As Long
    If Not PeriodContains(Period) Then Err.Raise AppErrInvalidRange, , "Period out of range."
    ToPeriodIndex = Period
End Function

Private Function MainStreamGradeContains(ByVal Grade As Long) As Boolean
    MainStreamGradeContains = (0 < Grade) And (Grade <= This.MainStreamGradeCount)
End Function

Private Function MainStreamClassContains(ByVal ClassNo As Long) As Boolean
    MainStreamClassContains = (0 < ClassNo) And (ClassNo <= This.MainStreamClassCount)
End Function

Private Function SpecialGradeContains(ByVal Grade As Long) As Boolean
    SpecialGradeContains = (0 < Grade) And (Grade <= This.SpecialGradeCount)
End Function

Private Function SpecialClassContains(ByVal ClassNo As Long) As Boolean
    SpecialClassContains = (0 < ClassNo) And (ClassNo <= This.SpecialClassCount)
End Function

Private Function PeriodContains(ByVal Period As Long) As Boolean
    PeriodContains = (0 < Period) And (Period <= This.PeriodCount)
End Function
