VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_TimeTableFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.EntityFactory")
Option Explicit

Private Type Member
    Mapper As Inf_TimeTableIDMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Mapper = New Inf_TimeTableIDMapper
End Sub

Public Function Create(ByVal Row As Inf_ClassHourRow) As Dom_TimeTable
    Dim ID As Dom_TimeTableID
    Set ID = This.Mapper.ToDomain(Row.RawID)
    If ID Is Nothing Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = Dom_ValueObjectFactory.CreateDateID(Row.NormDate)
    Dim Buffer As Long
    Buffer = Inf_NumericUtility.ExtractNumber(Row.RawValue)
    Dim Value As Dom_TimeTableValue
    Set Value = Dom_ValueObjectFactory.CreateTimeTableValue(Buffer)
    Dim Entry As Dom_TimeTable
    Set Entry = New Dom_TimeTable
    Entry.Initialize DateID, ID, Value
    Set Create = Entry
End Function
