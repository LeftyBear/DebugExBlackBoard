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
    If MainStreamGradeCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "", "MainStreamGradeCount must be positive number"
    If MainStreamClassCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "", "MainStreamClassCount must be positive number"
    If SpecialGradeCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "", "SpecialGradeCount must be positive number"
    If SpecialClassCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "", "SpecialClassCount must be positive number"
    If PeriodCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "", "PeriodCount must be positive number"
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
