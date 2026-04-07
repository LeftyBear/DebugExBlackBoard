VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    ID      As Dom_ScheduleID
    Value   As String
End Type

Private This As Member

Public Sub Initialize(ByVal ID As Dom_ScheduleID, ByVal Value As String)
    Set This.ID = ID
    This.Value = Value
End Sub

Friend Sub ChangeValueInternal(ByVal NewValue As String)
    This.Value = NewValue
End Sub

Public Function GetID() As Dom_ScheduleID
    Set GetID = This.ID
End Function

Public Function GetValue() As String
    GetValue = This.Value
End Function
