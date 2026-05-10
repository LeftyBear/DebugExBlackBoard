VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_YearlyEnrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.DTO")
Option Explicit
Private Type Member
    Aggregates As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Aggregates = New VBA.Collection
End Sub

Public Sub Add(ByVal Aggregate As Dom_DailyEnrollment)
    This.Aggregates.Add Aggregate
End Sub

Public Function Count() As Long
    Count = This.Aggregates.Count
End Function

Public Function Item(ByVal Index As Long) As Dom_DailyEnrollment
    Set Item = This.Aggregates.Item(Index)
End Function
