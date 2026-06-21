VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_TableViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.ViewModel")
Option Explicit

Private Type Member
    Table() As Variant
End Type

Private This As Member

Public Property Let Table(ByRef Table() As Variant)
    This.Table = Table
End Property

Public Property Get Table() As Variant()
    Table = This.Table
End Property
