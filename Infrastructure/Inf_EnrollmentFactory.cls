VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EnrollmentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.EntityFactory")
Option Explicit

Private Type Member
    Mapper As Inf_EnrollmentIDMapper
End Type

Private This As Member

Public Sub Inject(ByVal Mapper As Inf_EnrollmentIDMapper)
    Set This.Mapper = Mapper
End Sub

Public Function Create(ByVal Row As Inf_EnrollmentRow) As Dom_Enrollment
    Dim ID As Dom_EnrollmentID
    Set ID = This.Mapper.ToDomain(Row.RawID)
    If Row.RawTransfer <> vbNullString Then
        Dim State As Boolean
        State = Inf_BooleanUtility.ToBoolean(Row.RawTransfer)
        Dim Transfer As Dom_Transfer
        Set Transfer = Dom_ValueObjectFactory.CreateTransfer(State)
    End If
    If Row.Remarks <> vbNullString Then
        Dim Remarks As Dom_Remarks
        Set Remarks = Dom_ValueObjectFactory.CreateRemarks(Row.Remarks)
    End If
    If ID Is Nothing And Transfer Is Nothing Then Exit Function
    If ID Is Nothing And Remarks Is Nothing Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = Dom_ValueObjectFactory.CreateDateID(Row.NormDate)
    Dim Value As Dom_EnrollmentValue
    Set Value = Dom_ValueObjectFactory.CreateEnrollmentValue(Row.Value)
    Dim Entry As Dom_Enrollment
    Set Entry = New Dom_Enrollment
    Entry.Initialize DateID, ID, Value, Transfer, Remarks
    Set Create = Entry
End Function
