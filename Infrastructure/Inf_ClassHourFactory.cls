VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.EntityFactory")
Option Explicit

Private Type Member
    Mapper As Inf_ClassHourIDMapper
End Type

Private This As Member

Public Sub Inject(ByVal Mapper As Inf_ClassHourIDMapper)
    Set This.Mapper = Mapper
End Sub

Public Function Create(ByVal Row As Inf_ClassHourRow) As Dom_ClassHour
    Dim ID As Dom_ClassHourID
    Set ID = This.Mapper.ToDomain(Row.RawID)
    If ID Is Nothing Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = Dom_ValueObjectFactory.CreateDateID(Row.NormDate)
    Dim Buffer As Long
    Buffer = Inf_NumericUtility.ExtractNumber(Row.RawValue)
    Dim Value As Dom_ClassHourValue
    Set Value = Dom_ValueObjectFactory.CreateClassHourValue(Buffer)
    Dim Entry As Dom_ClassHour
    Set Entry = New Dom_ClassHour
    Entry.Initialize DateID, ID, Value
    Set Create = Entry
End Function
