VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourYearAggFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal ColumnGroup As VBA.Collection) As Dom_ClassHourYearAggregate
    Dim Yearly As Dom_ClassHourYearAggregate
    Set Yearly = New Dom_ClassHourYearAggregate
    Dim Daily As Dom_ClassHourAggregate
    For Each Daily In ColumnGroup
        Yearly.Add Daily
    Next
    Set Create = Yearly
End Function
