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

Friend Sub InitializeInternal(ByVal MainStreamGradeCount As Long, ByVal MainStreamClassCount As Long, ByVal PeriodCount As Long)
    This.MainStreamGradeCount = MainStreamGradeCount
    This.MainStreamClassCount = MainStreamClassCount
    This.PeriodCount = PeriodCount
End Sub

Public Function Create(ByVal MainStreamGradeCount As Long, ByVal MainStreamClassCount As Long, ByVal PeriodCount As Long) As Dom_UpperValue
    If MainStreamGradeCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "Dom_UpperValue", "学年は正の数である必要があります。"
    If MainStreamClassCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "Dom_UpperValue", "学級は正の数である必要があります。"
    If PeriodCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "Dom_UpperValue", "校時は正の数である必要があります。"
    Dim Result As Dom_UpperValue
    Set Result = New Dom_UpperValue
    Result.InitializeInternal MainStreamGradeCount, MainStreamClassCount, PeriodCount
    Set Create = Result
End Function

Public Property Get MainStreamGradeCount() As Long
    MainStreamGradeCount = This.MainStreamGradeCount
End Property

Public Property Get MainStreamClassCount() As Long
    MainStreamClassCount = This.MainStreamClassCount
End Property

Public Property Get PeriodCount() As Long
    PeriodCount = This.PeriodCount
End Property
