VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourHeaderMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.DTO")
Option Explicit
Private Type Member
    MapList As Scripting.Dictionary
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.MapList = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal Key As String, ByVal Column As Inf_ClassHourColumn)
    This.MapList.Add Key, Column
End Sub

Public Function Count() As Long
    Count = This.MapList.Count
End Function

Public Function Exists(ByVal Key As String) As Boolean
    Exists = This.MapList.Exists(Key)
End Function

Public Function Item(ByVal Key As String) As Inf_ClassHourColumn
    Set Item = This.MapList.Item(Key)
End Function
