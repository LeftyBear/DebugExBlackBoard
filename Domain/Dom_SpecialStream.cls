VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SpecialStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Index   As Long
    Name    As String
End Type

Private This As Member

Public Sub Initialize(ByVal Index As Long, ByVal Name As String)
    This.Index = Index
    This.Name = Name
End Sub

Public Property Get Index() As Long
    Index = This.Index
End Property

Public Property Get Name() As String
    Name = This.Name
End Property
