VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Period"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Value As Long
End Type

Private This As Member

Public Function Create(ByVal Value As Long) As Dom_Period
    If Value < 1 Then Err.Raise Dom_ErrNum.DomErrInvalidValue, "Dom_Period", "ZŽž‚Í³‚Ì”‚Å‚ ‚é•K—v‚ª‚ ‚è‚Ü‚·B"
    This.Value = Value
    Set Create = Me
End Function

Public Property Get Value() As Long
    Value = This.Value
End Property
