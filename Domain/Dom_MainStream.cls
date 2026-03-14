VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_MainStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    GradeCount  As Long
    ClassCount  As Long
End Type

Private This As Member

Public Sub Initialize(ByVal GradeCount As Long, ByVal ClassCount As Long)
    If GradeCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "Dom_MainStream", "学年は正の数である必要があります。"
    If ClassCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "Dom_MainStream", "学級は正の数である必要があります。"
    This.GradeCount = GradeCount
    This.ClassCount = ClassCount
End Sub

Public Property Get GradeCount() As Long
    GradeCount = This.GradeCount
End Property

Public Property Get ClassCount() As Long
    ClassCount = This.ClassCount
End Property
