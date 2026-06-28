VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DailyEnrollments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    DateID  As Dom_DateID
    Entries As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Entries = New VBA.Collection
End Sub

Public Sub Add(ByVal Entry As Dom_Enrollment)
    If This.DateID Is Nothing Then Set This.DateID = Entry.DateID
    This.Entries.Add Entry
End Sub

Public Sub UpdateValue(ByVal ID As Dom_EnrollmentID, ByVal NewValue As Long)
    Dim Target As Dom_Enrollment
    Set Target = FindByID(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_DailyEnrollments", "Enrollmentオブジェクトが必要です。"
    Dim Value As Dom_EnrollmentValue
    Set Value = New Dom_EnrollmentValue
    Value.Initialize NewValue
    Target.ChangeValueInternal Value
End Sub

Public Sub UpdateTransfer(ByVal NewValue As Boolean)
    Dim Target As Dom_Enrollment
    Set Target = FindTransfer
    If Target Is Nothing Then Err.Raise DomErrUnmatch, "Dom_DailyEnrollments", "Enrollmentオブジェクトが必要です。"
    Dim Transfer As Dom_Transfer
    Set Transfer = New Dom_Transfer
    Transfer.Initialize NewValue
    Target.ChangeTransferInternal Transfer
End Sub

Public Sub UpdateRemarks(ByVal NewValue As String)
    Dim Target As Dom_Enrollment
    Set Target = FindRemarks
    If Target Is Nothing Then Err.Raise DomErrUnmatch, "Dom_DailyEnrollments", "Enrollmentオブジェクトが必要です。"
    Dim Remarks As Dom_Remarks
    Set Remarks = New Dom_Remarks
    Remarks.Initialize NewValue
    Target.ChangeRemarksInternal Remarks
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Function Count() As Long
    Count = This.Entries.Count
End Function

Public Function Item(ByVal INDEX As Long) As Dom_Enrollment
    Set Item = This.Entries.Item(INDEX)
End Function

Private Function FindByID(ByVal ID As Dom_EnrollmentID) As Dom_Enrollment
    Dim i As Long
    For i = 1 To This.Entries.Count
        Dim Entry As Dom_Enrollment
        Set Entry = This.Entries.Item(i)
        If Entry.ID.Equals(ID) Then
            Set FindByID = Entry: Exit Function
        End If
    Next
End Function

Private Function FindTransfer() As Dom_Enrollment
    Dim i As Long
    For i = 1 To This.Entries.Count
        Dim Entry As Dom_Enrollment
        Set Entry = This.Entries.Item(i)
        If Entry.HasTransfer Then
            Set FindTransfer = Entry: Exit Function
        End If
    Next
End Function

Private Function FindRemarks() As Dom_Enrollment
    Dim i As Long
    For i = 1 To This.Entries.Count
        Dim Entry As Dom_Enrollment
        Set Entry = This.Entries.Item(i)
        If Entry.HasRemarks Then
            Set FindRemarks = Entry: Exit Function
        End If
    Next
End Function
