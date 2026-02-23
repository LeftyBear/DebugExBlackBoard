VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolStructure"
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
    SubjectCount            As Long
    PeriodCount             As Long
End Type

Private This As Member

Public Sub Initialize(ByVal MainStreamGradeCount As Long, ByVal MainStreamClassCount As Long, ByVal SpecialGradeCount As Long, ByVal SpecialClassCount As Long, ByVal SubjectCount As Long, ByVal PeriodCount As Long)
    This.MainStreamGradeCount = MainStreamGradeCount
    This.MainStreamClassCount = MainStreamClassCount
    This.SpecialGradeCount = SpecialGradeCount
    This.SpecialClassCount = SpecialClassCount
    This.SubjectCount = SubjectCount
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

Public Property Get SubjectCount() As Long
    SubjectCount = This.SubjectCount
End Property

Public Property Get PeriodCount() As Long
    PeriodCount = This.PeriodCount
End Property

Public Function MainStreamGradeContains(ByVal Grade As Long) As Boolean
    MainStreamGradeContains = (0 < Grade) And (Grade <= This.MainStreamGradeCount)
End Function

Public Function MainStreamClassContains(ByVal ClassNo As Long) As Boolean
    MainStreamClassContains = (0 < ClassNo) And (ClassNo <= This.MainStreamClassCount)
End Function

Public Function SpecialGradeContains(ByVal Grade As Long) As Boolean
    SpecialGradeContains = (0 < Grade) And (Grade <= This.SpecialGradeCount)
End Function

Public Function SpecialClassContains(ByVal ClassNo As Long) As Boolean
    SpecialClassContains = (0 < ClassNo) And (ClassNo <= This.SpecialClassCount)
End Function

Public Function SubjectContains(ByVal Subject As Long) As Boolean
    SubjectContains = (0 < Subject) And (Subject <= This.SubjectCount)
End Function

Public Function PeriodContains(ByVal Period As Long) As Boolean
    PeriodContains = (0 < Period) And (Period <= This.PeriodCount)
End Function

Public Function ToMainStreamGradeIndex(ByVal Grade As Long) As Long
    If Not MainStreamGradeContains(Grade) Then Err.Raise errInvalidRange, , "Grade out of range."
    ToMainStreamGradeIndex = Grade
End Function

Public Function ToMainStreamClassIndex(ByVal ClassNo As Long) As Long
    If Not MainStreamClassContains(ClassNo) Then Err.Raise errInvalidRange, , "Class out of range."
    ToMainStreamClassIndex = ClassNo
End Function

Public Function ToSpecialGradeIndex(ByVal Grade As Long) As Long
    If Not SpecialGradeContains(Grade) Then Err.Raise errInvalidRange, , "Grade out of range."
    ToSpecialGradeIndex = Grade
End Function

Public Function ToSpecialClassIndex(ByVal ClassNo As Long) As Long
    If Not SpecialClassContains(ClassNo) Then Err.Raise errInvalidRange, , "Class out of range."
    ToSpecialClassIndex = ClassNo
End Function

Public Function ToSubjectIndex(ByVal Subject As Long) As Long
    If Not SubjectContains(Subject) Then Err.Raise errInvalidRange, , "Subject out of range."
    ToSubjectIndex = Subject
End Function

Public Function ToPeriodIndex(ByVal Period As Long) As Long
    If Not PeriodContains(Period) Then Err.Raise errInvalidRange, , "Period out of range."
    ToPeriodIndex = Period
End Function
