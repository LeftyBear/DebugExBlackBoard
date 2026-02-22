VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentSummary"
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

Public Sub Add(ByVal StreamType As Dom_StreamType, ByVal Grade As Long, ByVal ClassNo As Long, ByVal Gender As Dom_Gender, ByVal Count As Long)
    Dim Key As String
    Key = CStr(StreamType) & "_" & CStr(Grade) & "_" & CStr(ClassNo) & "_" & CStr(Gender)
    If Not This.TotalList.Exists(Key) Then
        This.TotalList.Add Key, 0
    End If
    This.TotalList(Key) = This.TotalList(Key) + Count
End Sub

Public Function GetTotal(ByVal StreamType As Dom_StreamType, ByVal Grade As Long, ByVal ClassNo As Long, ByVal Gender As Dom_Gender) As Long
    Dim Key As String
    Key = CStr(StreamType) & "_" & CStr(Grade) & "_" & CStr(ClassNo) & "_" & CStr(Gender)
    If This.TotalList.Exists(Key) Then
        GetTotal = This.TotalList(Key)
    End If
End Function
