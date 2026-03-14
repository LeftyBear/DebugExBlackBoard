VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    MainStream      As Dom_MainStream
    SpecialStream   As Dom_SpecialStream
    Subject         As Dom_Subject
    TimeTable       As Dom_TimeTable
End Type

Private This As Member

Public Sub Initialize(ByVal MainStream As Dom_MainStream, ByVal SpecialStream As Dom_SpecialStream, ByVal Subject As Dom_Subject, ByVal TimeTable As Dom_TimeTable)
    Set This.MainStream = MainStream
    Set This.SpecialStream = SpecialStream
    Set This.Subject = Subject
    Set This.TimeTable = TimeTable
End Sub

Public Property Get MainStreamGradeCount() As Long
    MainStreamGradeCount = This.MainStream.GradeCount
End Property

Public Property Get MainStreamClassCount() As Long
    MainStreamClassCount = This.MainStream.ClassCount
End Property

Public Property Get SpecialStreamGradeCount() As Long
    SpecialStreamGradeCount = 1
End Property

Public Property Get SpecialStreamClassCount() As Long
    SpecialStreamClassCount = This.SpecialStream.Count
End Property

Public Property Get SubjectCount() As Long
    SubjectCount = This.Subject.Count
End Property

Public Property Get PeriodCount() As Long
    PeriodCount = This.TimeTable.PeriodCount
End Property

Public Function SubjectName(ByVal Index As Long) As String
    If Not SubjectIndexContains(Index) Then Err.Raise DomErrInvalidRange, , "インデックスは次の範囲である必要があります。 " & "範囲: 1 ～ " & Me.SubjectCount
    SubjectName = This.Subject.Name(Index)
End Function

Public Function SubjectIndex(ByVal SubjectName As String) As Long
    SubjectIndex = This.Subject.Index(SubjectName)
End Function

Public Function SubjectTargetGrade(ByVal Index As Long) As String
    If Not SubjectIndexContains(Index) Then Err.Raise DomErrInvalidRange, , "インデックスは次の範囲である必要があります。 " & "範囲: 1 ～ " & Me.SubjectCount
    SubjectTargetGrade = This.Subject.TargetGrade(Index)
End Function

Public Function SpecialClassName(ByVal Index As Long) As String
    If Not SpecialClassIndexContains(Index) Then Err.Raise DomErrInvalidRange, , "インデックスは次の範囲である必要があります。 " & "範囲: 1 ～ " & Me.SpecialStreamClassCount
    SpecialClassName = This.SpecialStream.Name(Index)
End Function

Private Function SubjectIndexContains(ByVal Index As Long) As Boolean
    SubjectIndexContains = (0 < Index) And (Index <= Me.SubjectCount)
End Function

Private Function SpecialClassIndexContains(ByVal Index As Long) As Boolean
    SpecialClassIndexContains = (0 < Index) And (Index <= Me.SpecialStreamClassCount)
End Function
