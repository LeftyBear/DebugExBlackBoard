VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_Periods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.DTOCollection")
Option Explicit
Private Type Member
    Values As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Values = New VBA.Collection
End Sub

Public Sub Add(ByVal Grade As Long, ByVal Value As Long)
    This.Values.Add Value, Grade
End Sub

Public Function Count() As Long
    Count = This.Values.Count
End Function

Public Function Item(ByVal Grade As Long) As Long
    If This.Values.Item(Grade) Is Nothing Then Exit Function
    Item = This.Values.Item(Grade)
End Function
