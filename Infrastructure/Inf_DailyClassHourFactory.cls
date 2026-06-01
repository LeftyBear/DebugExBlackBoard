VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyClassHourFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.EntityFactory")
Option Explicit

Private Type Member
    IDMapper    As Inf_ClassHourIDMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.IDMapper = New Inf_ClassHourIDMapper
End Sub

Public Function Create(ByVal TargetDate As Date, ByVal Row As Inf_ClassHourRow) As Dom_ClassHour
    Dim ID As Dom_ClassHourID
    Set ID = This.IDMapper.ToDomain(Row.RawID)
    If ID Is Nothing Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = Dom_ValueObjectFactory.CreateDateID(TargetDate)
    Dim Value As Dom_ClassHourValue
    Set Value = Dom_ValueObjectFactory.CreateClassHourValue(Row.Value)
    Dim Entity As Dom_ClassHour
    Set Entity = New Dom_ClassHour
    Entity.Initialize DateID, ID, Value
    Set Create = Entity
End Function
