VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Index           As Long
    Name            As String
    TargetGrade     As String
    Mark            As String
End Type

Private This As Member

Public Sub Initialize(ByVal Index As Long, Name As String, ByVal TargetGrade As String, ByVal Mark As String)
    If Index < 1 Then Err.Raise DomErrInvalidValue, "Dom_Subject", "インデックスは正の数である必要があります。"
    This.Index = Index
    This.Name = Name
    This.TargetGrade = TargetGrade
    This.Mark = Mark
End Sub

Public Property Get Index() As Long
    Index = This.Index
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Get TargetGrade() As String
    TargetGrade = This.TargetGrade
End Property

Public Property Get Mark() As String
    Mark = This.Mark
End Property
