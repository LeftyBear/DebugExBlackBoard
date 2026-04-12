VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal ID As Dom_ClassHourID, ByVal RawValue As Long) As Dom_ClassHour
    Dim Value As Dom_ClassHourValue
    Set Value = Dom_ValueObjectFactory.CreateClassHourValue(RawValue)
    Dim Result As Dom_ClassHour
    Set Result = New Dom_ClassHour
    Result.Initialize ID, Value
    Set Create = Result
End Function
