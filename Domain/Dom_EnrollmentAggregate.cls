VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    DateID      As Dom_DateID
    Entities    As VBA.Collection
    Transfer    As Dom_Transfer
    Remarks     As Dom_Remarks
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Entities = New VBA.Collection
End Sub

Public Sub SetDateID(ByVal DateID As Dom_DateID)
    Set This.DateID = DateID
End Sub

Public Sub SetTransfer(ByVal Transfer As Dom_Transfer)
    Set This.Transfer = Transfer
End Sub

Public Sub SetRemarks(ByVal Remarks As Dom_Remarks)
    Set This.Remarks = Remarks
End Sub

Public Sub AddEnrollment(ByVal Entity As Dom_Enrollment)
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

Public Sub UpdateTransfer(ByVal DateID As Dom_DateID, ByVal NewValue As Boolean)
    If Not This.Transfer.DateID.Equals(DateID) Then Err.Raise DomErrUnmatch, "Dom_EnrollmentAggregate", "日付が一致する必要があります。"
    Dim Transfer As Dom_Transfer
    Set Transfer = New Dom_Transfer
    Transfer.Initialize DateID, NewValue
    Set This.Transfer = Transfer
End Sub

Public Sub UpdateRemarks(ByVal DateID As Dom_DateID, ByVal NewValue As String)
    If Not This.Remarks.DateID.Equals(DateID) Then Err.Raise DomErrUnmatch, "Dom_EnrollmentAggregate", "日付が一致する必要があります。"
    Dim Remarks As Dom_Remarks
    Set Remarks = New Dom_Remarks
    Remarks.Initialize DateID, NewValue
    Set This.Remarks = Remarks
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Property Get Transfer() As Dom_Transfer
    Set Transfer = This.Transfer
End Property

Public Property Get Remarks() As Dom_Remarks
    Set Remarks = This.Remarks
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
