VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ColumnDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.DTO")
Option Explicit

Private Type Member
    HeaderName  As String
    RawValue    As String
End Type

Private This As Member

Public Sub Initialize(ByVal HeaderName As String, ByVal RawValue As String)
    This.HeaderName = HeaderName
    This.RawValue = RawValue
End Sub

Public Property Get HeaderName() As String
    HeaderName = This.HeaderName
End Property

Public Property Get RawValue() As String
    RawValue = This.RawValue
End Property
