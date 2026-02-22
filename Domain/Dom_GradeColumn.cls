VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_GradeColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Number As Long
End Type

Private This As Member

Public Sub Initialize(ByVal Number As Long)
    This.Number = Number
End Sub

Public Property Get Number() As Long
    Number = This.Number
End Property
