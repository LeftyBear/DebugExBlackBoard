VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.EntityFactory")
Option Explicit

Private Type Member
    ValueObjectFactory As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal ValueObjectFactory As Dom_ValueObjectFactory)
    Set This.ValueObjectFactory = ValueObjectFactory
End Sub

Public Function Create(ByVal Row As App_EnrollmentRow) As Dom_Enrollment
    Dim DateID As Dom_EnrollmentID
    Set DateID = This.ValueObjectFactory.CreateDateID(Row.NormDate)
    Dim ID As Dom_EnrollmentID
    Set ID = CreateID(Row)
    Dim Value As Dom_EnrollmentValue
    Set Value = This.ValueObjectFactory.CreateEnrollmentValue(Row.Value)
    If Row.RawTransfer <> vbNullString Then
        Dim Transfer As Dom_Transfer
        Set Transfer = This.ValueObjectFactory.CreateTransfer(Row.RawTransfer)
    End If
    If Row.RawRemarks <> vbNullString Then
        Dim Remarks As Dom_Remarks
        Set Remarks = This.ValueObjectFactory.CreateRemarks(Row.RawRemarks)
    End If
    If ID Is Nothing And Transfer Is Nothing And Remarks Is Nothing Then Exit Function
    Dim Entity As Dom_Enrollment
    Set Entity = New Dom_Enrollment
    Entity.Initialize DateID, ID, Value, Transfer, Remarks
    Set Create = Entity
End Function

Private Function CreateID(ByVal Row As App_EnrollmentRow) As Dom_EnrollmentID
    Dim StreamType As Dom_StreamType
    Set StreamType = This.ValueObjectFactory.CreateStreamType(Row.TypeCode)
    If StreamType.Value = 0 Then Exit Function
    Dim Grade As Dom_Grade
    Set Grade = This.ValueObjectFactory.CreateGrade(Row.Grade)
    If Grade.Value = 0 Then Exit Function
    Dim ClassNo As Dom_ClassNo
    Set ClassNo = This.ValueObjectFactory.CreateClassNo(Row.ClassNo)
    If ClassNo.Value = 0 Then Exit Function
    Dim Gender As Dom_Gender
    Set Gender = This.ValueObjectFactory.CreateGender(Row.Gender)
    If Gender.Value = 0 Then Exit Function
    Set CreateID = This.ValueObjectFactory.CreateEnrollmentID(StreamType, Grade, ClassNo, Gender)
End Function
