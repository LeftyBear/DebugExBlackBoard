VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    ID      As Dom_TimeTableID
    Mark    As Dom_TimeTableMark
End Type

Private This As Member

Public Sub Initialize(ByVal ID As Dom_TimeTableID, ByVal Mark As Dom_TimeTableMark)
    Set This.ID = ID
    Set This.Mark = Mark
End Sub

Friend Sub ChangeValueInternal(ByVal NewMark As Dom_TimeTableMark)
    Set This.Mark = NewMark
End Sub

Public Function GetID() As Dom_TimeTableID
    Set GetID = This.ID
End Function

Public Function GetMark() As String
    GetMark = This.Mark.Value
End Function
