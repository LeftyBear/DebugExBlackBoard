VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    ID      As Dom_ClassHourID
    Value   As Dom_ClassHourValue
End Type

Private This As Member

Public Sub Initialize(ByVal ID As Dom_ClassHourID, ByVal Value As Dom_ClassHourValue)
    Set This.ID = ID
    Set This.Value = Value
End Sub

Friend Sub ChangeValueInternal(ByVal NewValue As Dom_ClassHourValue)
    Set This.Value = NewValue
End Sub

Public Function GetID() As Dom_ClassHourID
    Set GetID = This.ID
End Function

Public Function GetValue() As Long
    GetValue = This.Value.Value
End Function
