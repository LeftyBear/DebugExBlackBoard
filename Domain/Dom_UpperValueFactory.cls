VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_UpperValueFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal Row As App_UpperValueRow) As Dom_UpperValue
    Dim Entity As Dom_UpperValue
    Set Entity = New Dom_UpperValue
    Entity.Initialize Row.MainStreamGradeCount, Row.MainStreamClassCount, Row.PeriodCount
    Set Create = Entity
End Function
