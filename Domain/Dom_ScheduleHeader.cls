VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.DTO")
Option Explicit

Private Type Member
    RawDate As String
    RawText As String
End Type

Private This As Member

Public Sub Initialize(ByVal RawDate As String, ByVal RawText As String)
    This.RawDate = RawDate
    This.RawText = RawText
End Sub

Public Property Get RawDate() As String
    RawDate = This.RawDate
End Property

Public Property Get RawText() As String
    RawText = This.RawText
End Property
