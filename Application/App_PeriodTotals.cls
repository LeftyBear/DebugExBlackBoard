VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_PeriodTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.DTOCollection")
Option Explicit
Private Type Member
    Totals As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Totals = New VBA.Collection
End Sub

Public Sub Add(ByVal Grade As Long, ByVal Value As Long)
    This.Totals.Add Value, Grade
End Sub

Public Function Count() As Long
    Count = This.Totals.Count
End Function

Public Function Item(ByVal Grade As Long) As Long
    If This.Totals.Item(Grade) Is Nothing Then Exit Function
    Item = This.Totals.Item(Grade)
End Function
