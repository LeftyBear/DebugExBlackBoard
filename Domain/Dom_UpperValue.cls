VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_UpperValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    MainStreamGradeCount    As Long
    MainStreamClassCount    As Long
    PeriodCount             As Long
End Type

Private This As Member

Public Sub Initialize(ByVal MainStreamGradeCount As Long, ByVal MainStreamClassCount As Long, ByVal PeriodCount As Long)
    If MainStreamGradeCount < 1 Then Err.Raise DomErrInvalidValue, "Dom_UpperValue", "学年は正の数である必要があります。"
    If MainStreamClassCount < 1 Then Err.Raise DomErrInvalidValue, "Dom_UpperValue", "学級は正の数である必要があります。"
    If PeriodCount < 1 Then Err.Raise DomErrInvalidValue, "Dom_UpperValue", "校時は正の数である必要があります。"
    This.MainStreamGradeCount = MainStreamGradeCount
    This.MainStreamClassCount = MainStreamClassCount
    This.PeriodCount = PeriodCount
End Sub

Public Property Get MainStreamGradeCount() As String
    MainStreamGradeCount = This.MainStreamGradeCount
End Property

Public Property Get MainStreamClassCount() As String
    MainStreamClassCount = This.MainStreamClassCount
End Property

Public Property Get PeriodCount() As String
    PeriodCount = This.PeriodCount
End Property

Public Function Equals(ByVal Other As Dom_UpperValue) As Boolean
    If This.MainStreamGradeCount <> Other.MainStreamGradeCount Then Exit Function
    If This.MainStreamClassCount <> Other.MainStreamClassCount Then Exit Function
    If This.PeriodCount <> Other.PeriodCount Then Exit Function
    Equals = True
End Function
