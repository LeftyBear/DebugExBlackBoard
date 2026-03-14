VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    PeriodCount As Long
End Type

Private This As Member

Public Sub Initialize(ByVal PeriodCount As Long)
    If PeriodCount < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "Dom_TimeTable", "ŽžŠÔŠ„”‚Í³‚Ì”‚Å‚ ‚é•K—v‚ª‚ ‚è‚Ü‚·B"
    This.PeriodCount = PeriodCount
End Sub

Public Property Get PeriodCount() As Long
    PeriodCount = This.PeriodCount
End Property
