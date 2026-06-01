VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyEnrollmentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.EntityFactory")
Option Explicit

Private Type Member
    IDMapper        As Inf_EnrollmentIDMapper
    TransferMapper  As Inf_TransferMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.IDMapper = New Inf_EnrollmentIDMapper
    Set This.TransferMapper = New Inf_TransferMapper
End Sub

Public Function Create(ByVal TargetDate As Date, ByVal Row As Inf_EnrollmentRow) As Dom_Enrollment
    Dim ID As Dom_EnrollmentID
    Set ID = This.IDMapper.ToDomain(Row.RawID)
    If Row.RawTransfer <> vbNullString Then
        Dim Transfer As Dom_Transfer
        Set Transfer = This.TransferMapper.ToDomain(Row.RawTransfer)
    End If
    If Row.Remarks <> vbNullString Then
        Dim Remarks As Dom_Remarks
        Set Remarks = Dom_ValueObjectFactory.CreateRemarks(Row.Remarks)
    End If
    If ID Is Nothing And Transfer Is Nothing Then Exit Function
    If ID Is Nothing And Remarks Is Nothing Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = Dom_ValueObjectFactory.CreateDateID(TargetDate)
    Dim Value As Dom_EnrollmentValue
    Set Value = Dom_ValueObjectFactory.CreateEnrollmentValue(Row.Value)
    Dim Entity As Dom_Enrollment
    Set Entity = New Dom_Enrollment
    Entity.Initialize DateID, ID, Value, Transfer, Remarks
    Set Create = Entity
End Function
