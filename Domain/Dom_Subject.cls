VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Name        As Dom_SubjectName
    TargetGrade As Dom_SubjectTargetGrade
    Mark        As Dom_SubjectMark
End Type

Private This As Member

Public Sub Initialize(ByVal Name As Dom_SubjectName, ByVal TargetGrade As Dom_SubjectTargetGrade, ByVal Mark As Dom_SubjectMark)
    Set This.Name = Name
    Set This.TargetGrade = TargetGrade
    Set This.Mark = Mark
End Sub

Public Property Get Name() As Dom_SubjectName
    Set Name = This.Name
End Property

Public Property Get TargetGrade() As Dom_SubjectTargetGrade
    Set TargetGrade = This.TargetGrade
End Property

Public Property Get Mark() As Dom_SubjectMark
    Set Mark = This.Mark
End Property

Public Function Equals(ByVal Other As Dom_Subject) As Boolean
    If Not This.Name.Equals(Other.Name) Then Exit Function
    If Not This.TargetGrade.Equals(Other.TargetGrade) Then Exit Function
    If Not This.Mark.Equals(Other.Mark) Then Exit Function
    Equals = True
End Function
