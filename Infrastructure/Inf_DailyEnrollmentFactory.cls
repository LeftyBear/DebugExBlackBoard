VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyEnrollmentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.EntityFactory")
Option Explicit

Public Function Create(ByVal TargetDate As Date, ByVal Row As Inf_EnrollmentRow) As Dom_Enrollment
    Dim Mapper As New Inf_EnrollmentIDMapper
    Dim ID As Dom_EnrollmentID
    Set ID = Mapper.ToDomain(Row.RawID)
    If Row.RawTransfer <> vbNullString Then
        Dim State As Boolean
        State = App_BooleanUtility.NormalizeToBoolean(Row.RawTransfer)
        Dim Transfer As Dom_Transfer
        Dim Factory As New Dom_ValueObjectFactory
        Set Transfer = Factory.CreateTransfer(State)
    End If
    If Row.Remarks <> vbNullString Then
        Dim Remarks As Dom_Remarks
        Set Remarks = Factory.CreateRemarks(Row.Remarks)
    End If
    If ID Is Nothing And Transfer Is Nothing Then Exit Function
    If ID Is Nothing And Remarks Is Nothing Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = Factory.CreateDateID(TargetDate)
    Dim Value As Dom_EnrollmentValue
    Set Value = Factory.CreateEnrollmentValue(Row.Value)
    Dim Entity As Dom_Enrollment
    Set Entity = New Dom_Enrollment
    Entity.Initialize DateID, ID, Value, Transfer, Remarks
    Set Create = Entity
End Function
