VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyClassHourFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.EntityFactory")
Option Explicit

Private Type Member
    Factory         As Dom_ValueObjectFactory
    IDMapper        As Inf_ClassHourIDMapper
    TransferMapper  As Inf_TransferMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Factory = New Dom_ValueObjectFactory
    Set This.IDMapper = New Inf_ClassHourIDMapper
    Set This.TransferMapper = New Inf_TransferMapper
    This.IDMapper.Initialize This.Factory
    This.TransferMapper.Initialize This.Factory
End Sub

Public Function Create(ByVal TargetDate As Date, ByVal Row As Inf_ClassHourRow) As Dom_ClassHour
    Dim ID As Dom_ClassHourID
    Set ID = This.IDMapper.ToDomain(Row.RawID)
    If Row.RawTransfer <> vbNullString Then
        Dim Transfer As Dom_Transfer
        Set Transfer = This.TransferMapper.ToDomain(Row.RawTransfer)
    End If
    If Row.Remarks <> vbNullString Then
        Dim Remarks As Dom_Remarks
        Set Remarks = This.Factory.CreateRemarks(Row.Remarks)
    End If
    If ID Is Nothing And Transfer Is Nothing Then Exit Function
    If ID Is Nothing And Remarks Is Nothing Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = This.Factory.CreateDateID(TargetDate)
    Dim Value As Dom_ClassHourValue
    Set Value = This.Factory.CreateClassHourValue(Row.Value)
    Dim Entity As Dom_ClassHour
    Set Entity = New Dom_ClassHour
    Entity.Initialize DateID, ID, Value, Transfer, Remarks
    Set Create = Entity
End Function
