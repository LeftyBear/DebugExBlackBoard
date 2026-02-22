VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScalarColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Code As Long
End Type

Private This As Member

Public Sub Initialize(ByVal Code As Long)
    This.Code = Code
End Sub

Public Property Get Code() As Long
    Code = This.Code
End Property
