VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    DateID      As Dom_DateID
    Entities    As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Entities = New VBA.Collection
End Sub

Public Sub Add(ByVal Entity As Dom_Enrollment)
    If This.DateID Is Nothing Then Set This.DateID = Entity.DateID
    This.Entities.Add Entity
End Sub

Public Sub UpdateValue(ByVal ID As Dom_EnrollmentID, ByVal NewValue As Long)
    Dim Target As Dom_Enrollment
    Set Target = FindByID(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_EnrollmentAggregate", "Enrollmentオブジェクトが必要です。"
    Dim Value As Dom_EnrollmentValue
    Set Value = New Dom_EnrollmentValue
    Value.Initialize NewValue
    Target.ChangeValueInternal Value
End Sub

Public Sub UpdateTransfer(ByVal NewValue As Boolean)
    Dim Target As Dom_Enrollment
    Set Target = FindTransfer
    If Target Is Nothing Then Err.Raise DomErrUnmatch, "Dom_EnrollmentAggregate", "Enrollmentオブジェクトが必要です。"
    Dim Transfer As Dom_Transfer
    Set Transfer = New Dom_Transfer
    Transfer.Initialize NewValue
    Target.ChangeTransferInternal Transfer
End Sub

Public Sub UpdateRemarks(ByVal NewValue As String)
    Dim Target As Dom_Enrollment
    Set Target = FindRemarks
    If Target Is Nothing Then Err.Raise DomErrUnmatch, "Dom_EnrollmentAggregate", "Enrollmentオブジェクトが必要です。"
    Dim Remarks As Dom_Remarks
    Set Remarks = New Dom_Remarks
    Remarks.Initialize NewValue
    Target.ChangeRemarksInternal Remarks
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Function Count() As Long
    Count = This.Entities.Count
End Function

Public Function Item(ByVal Index As Long) As Dom_Enrollment
    Set Item = This.Entities.Item(Index)
End Function

Private Function FindByID(ByVal ID As Dom_EnrollmentID) As Dom_Enrollment
    Dim i As Long
    For i = 1 To This.Entities.Count
        Dim Entity As Dom_Enrollment
        Set Entity = This.Entities.Item(i)
        If Entity.ID.Equals(ID) Then
            Set FindByID = Entity: Exit Function
        End If
    Next
End Function

Private Function FindTransfer() As Dom_Enrollment
    Dim i As Long
    For i = 1 To This.Entities.Count
        Dim Entity As Dom_Enrollment
        Set Entity = This.Entities.Item(i)
        If Entity.HasTransfer Then
            Set FindTransfer = Entity: Exit Function
        End If
    Next
End Function

Private Function FindRemarks() As Dom_Enrollment
    Dim i As Long
    For i = 1 To This.Entities.Count
        Dim Entity As Dom_Enrollment
        Set Entity = This.Entities.Item(i)
        If Entity.HasRemarks Then
            Set FindRemarks = Entity: Exit Function
        End If
    Next
End Function
