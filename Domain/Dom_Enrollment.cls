VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Enrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    ID          As Dom_EnrollmentID
    NormDate    As Date
    Value       As Dom_EnrollmentValue
    Transfer    As Boolean
    Remarks     As String
End Type

Private This As Member

Public Sub Initialize(ByVal ID As Dom_EnrollmentID, ByVal NormDate As Date, ByVal Value As Dom_EnrollmentValue, ByVal Transfer As Boolean, ByVal Remarks As String)
    Set This.ID = ID
    This.NormDate = NormDate
    Set This.Value = Value
    This.Transfer = Transfer
    This.Remarks = Remarks
End Sub

Friend Sub ChangeValueInternal(ByVal NewValue As Dom_EnrollmentValue)
    Set This.Value = NewValue
End Sub

Friend Sub ChangeTransferInternal(ByVal NewState As Boolean)
    This.Transfer = NewState
End Sub

Friend Sub ChangeRemarksInternal(ByVal NewText As String)
    This.Remarks = NewText
End Sub

Public Function GetID() As Dom_EnrollmentID
    Set GetID = This.ID
End Function

Public Function GetDate() As Date
    GetDate = This.NormDate
End Function

Public Function GetValue() As Dom_EnrollmentValue
    Set GetValue = This.Value
End Function

Public Function GetTransfer() As Boolean
    GetTransfer = This.Transfer
End Function

Public Function GetRemarks() As String
    GetRemarks = This.Remarks
End Function
