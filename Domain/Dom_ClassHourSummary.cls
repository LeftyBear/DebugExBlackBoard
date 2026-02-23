VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    TotalList As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize()
    Set This.TotalList = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal Grade As Long, ByVal ClassNo As Long, ByVal Gender As Long, ByVal Count As Long)
    Dim Key As String
    Key = CStr(Grade) & "_" & CStr(ClassNo) & "_" & CStr(Gender)
    If Not This.TotalList.Exists(Key) Then
        This.TotalList.Add Key, 0
    End If
    This.TotalList.Item(Key) = This.TotalList.Item(Key) + Count
End Sub

Public Function GetTotal(ByVal Grade As Long, ByVal ClassNo As Long, ByVal Gender As Long) As Long
    Dim Key As String
    Key = CStr(Grade) & "_" & CStr(ClassNo) & "_" & CStr(Gender)
    If This.TotalList.Exists(Key) Then
        GetTotal = This.TotalList.Item(Key)
    End If
End Function

