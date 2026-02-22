VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SchoolYearKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Application.Key

Option Explicit

Private Type Member
    SchoolYear As Long
End Type

Private This As Member

Public Sub Initialize(ByVal SchoolYear As Long)
    If Not (0 < SchoolYear) Then
        Err.Raise App_ErrNum.NotPositiveNumber, TypeName(Me), "SchoolYear must be positive."
    End If
    This.SchoolYear = SchoolYear
End Sub

Public Property Get Value() As Long
    Value = This.SchoolYear
End Property

