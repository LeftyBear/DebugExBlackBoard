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
    ID          As Dom_ClassHourID
    NormDate    As Date
    Hours       As Dom_ClassHourValue
    Mark        As String
End Type

Private This As Member

Public Sub Initialize(ByVal ID As Dom_ClassHourID, ByVal NormDate As Date, ByVal Hours As Dom_ClassHourValue, ByVal Mark As String)
    Set This.ID = ID
    This.NormDate = NormDate
    Set This.Hours = Hours
    This.Mark = Mark
End Sub

Friend Sub ChangeValueInternal(ByVal NewValue As Dom_ClassHourValue)
    Set This.Hours = NewValue
End Sub

Friend Sub ChangeMarkInternal(ByVal NewMark As String)
    This.Mark = NewMark
End Sub

Public Function GetID() As Dom_ClassHourID
    Set GetID = This.ID
End Function

Public Function GetDate() As Date
    GetDate = This.NormDate
End Function

Public Function GetHours() As Dom_ClassHourValue
    Set GetHours = This.Hours
End Function

Public Function GetMark() As String
    GetMark = This.Mark
End Function
