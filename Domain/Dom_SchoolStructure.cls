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
    PeriodCount             As Long
    SubjectName             As Dom_SubjectName
    SubjectTargetGrade      As Dom_SubjectTargetGrade
End Type

Private This As Member

Public Sub Initialize(ByVal MainStreamGradeCount As Long, ByVal MainStreamClassCount As Long, ByVal SpecialGradeCount As Long, ByVal SpecialClassCount As Long, ByVal PeriodCount As Long, ByVal SubjectName As Dom_SubjectName, ByVal SubjectTargetGrade As Dom_SubjectTargetGrade)
    This.MainStreamGradeCount = MainStreamGradeCount
    This.MainStreamClassCount = MainStreamClassCount
    This.SpecialGradeCount = SpecialGradeCount
    This.SpecialClassCount = SpecialClassCount
    This.PeriodCount = PeriodCount
    This.SubjectName = SubjectName
    This.SubjectTargetGrade = SubjectTargetGrade
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

Public Property Get SubjectCount() As Long
    SubjectCount = This.SubjectName.Count
End Property

Public Property Get SubjectName(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise AppErrInvalidRange, , "Index out of range"
    SubjectName = This.SubjectName.GetValue(Index)
End Property

Public Property Get SubjectTargetGrade(ByVal Index As Long) As String
    If Not IndexContains(Index) Then Err.Raise AppErrInvalidRange, , "Index out of range"
    SubjectTargetGrade = This.SubjectTargetGrade.GetValue(Index)
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

Public Function ToSubjectIndex(ByVal TargetName As String) As Long
    Dim Index As Long
    Index = GetIndex(TargetName)
    If Index = 0 Then Err.Raise AppErrNotFoundItem, , "Not found item: " & TargetName
    ToSubjectIndex = Index
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

Private Function IndexContains(ByVal Index As Long) As Boolean
    IndexContains = (0 < Index) And (Index <= This.SubjectName.Count)
End Function

Private Function SubjectContains(ByVal TargetName As String) As Boolean
    Dim i As Long
    For i = 1 To This.SubjectName.Count
        If This.SubjectName.GetValue(i) Like TargetName Then SubjectContains = True
    Next
End Function

Private Function GetIndex(ByVal TargetName As String) As Long
    Dim i As Long
    For i = 1 To This.SubjectName.Count
        If This.SubjectName.GetValue(i) Like TargetName Then GetIndex = i
    Next
End Function
