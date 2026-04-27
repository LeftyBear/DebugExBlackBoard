VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourYearAggFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Private Type Member
    Structure As Dom_SchoolStructure
    CHFactory As App_ClassHourFactory
    TTFactory As App_TimeTableFactory
    VOFactory As App_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourFactory As App_ClassHourFactory, ByVal TimeTableFactory As App_TimeTableFactory, ByVal VOFactory As App_ValueObjectFactory)
    Set This.CHFactory = ClassHourFactory
    Set This.TTFactory = TimeTableFactory
    Set This.VOFactory = VOFactory
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Create(ByVal Rows As VBA.Collection) As Dom_ClassHourYearAggregate
    Dim Result As Dom_ClassHourYearAggregate
    Set Result = New Dom_ClassHourYearAggregate
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Aggregate As Dom_ClassHourAggregate
        Result.Add Aggregate
    Next
    Set Create = Result
End Function
