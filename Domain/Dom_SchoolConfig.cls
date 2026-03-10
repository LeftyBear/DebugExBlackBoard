VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    LimitValue  As Dom_LimitValue
    Subject     As Dom_Subject
End Type

Private This As Member

Public Sub Initialize(ByVal LimitValue As Dom_LimitValue, ByVal Subject As Dom_Subject)
    Set This.LimitValue = LimitValue
    Set This.Subject = Subject
End Sub

Public Property Get MainStreamGradeCount() As Long
    MainStreamGradeCount = This.LimitValue.MainStreamGradeCount
End Property

Public Property Get MainStreamClassCount() As Long
    MainStreamClassCount = This.LimitValue.MainStreamClassCount
End Property

Public Property Get SpecialGradeCount() As Long
    SpecialGradeCount = This.LimitValue.SpecialGradeCount
End Property

Public Property Get SpecialClassCount() As Long
    SpecialClassCount = This.LimitValue.SpecialClassCount
End Property

Public Property Get PeriodCount() As Long
    PeriodCount = This.LimitValue.PeriodCount
End Property

Public Property Get SubjectCount() As Long
    SubjectCount = This.Subject.Count
End Property

Public Property Get SubjectName(ByVal Index As Long) As String
    If Not SubjectIndexContains(Index) Then Err.Raise AppErrInvalidRange, , "Index out of range"
    SubjectName = This.Subject.Name(Index)
End Property

Public Property Get SubjectTargetGrade(ByVal Index As Long) As String
    If Not SubjectIndexContains(Index) Then Err.Raise AppErrInvalidRange, , "Index out of range"
    SubjectTargetGrade = This.Subject.TargetGrade(Index)
End Property

Private Function SubjectIndexContains(ByVal Index As Long) As Boolean
    SubjectIndexContains = (0 < Index) And (Index <= This.Subject.Count)
End Function
