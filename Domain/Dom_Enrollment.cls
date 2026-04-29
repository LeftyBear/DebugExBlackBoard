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
    DateID      As Dom_DateID
    ID          As Dom_EnrollmentID
    Value       As Dom_EnrollmentValue
    Transfer    As Dom_Transfer
    Remarks     As Dom_Remarks
End Type

Private This As Member

Public Sub Initialize(ByVal DateID As Dom_DateID, ByVal ID As Dom_EnrollmentID, ByVal Value As Dom_EnrollmentValue, Optional ByVal Transfer As Dom_Transfer, Optional ByVal Remarks As Dom_Remarks)
    Set This.DateID = DateID
    Set This.ID = ID
    Set This.Value = Value
    Set This.Transfer = Transfer
    Set This.Remarks = Remarks
End Sub

Friend Sub ChangeValueInternal(ByVal NewValue As Dom_EnrollmentValue)
    Set This.Value = NewValue
End Sub

Friend Sub ChangeTransferInternal(ByVal NewValue As Dom_Transfer)
    Set This.Transfer = NewValue
End Sub

Friend Sub ChangeRemarksInternal(ByVal NewValue As Dom_Remarks)
    Set This.Remarks = NewValue
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Property Get ID() As Dom_EnrollmentID
    Set ID = This.ID
End Property

Public Property Get Value() As Dom_EnrollmentValue
    Set Value = This.Value
End Property

Public Property Get Transfer() As Dom_Transfer
    Set Transfer = This.Transfer
End Property

Public Property Get Remarks() As Dom_Remarks
    Set Remarks = This.Remarks
End Property

Public Function HasTransfer() As Boolean
    If This.Transfer Is Nothing Then Exit Function
    HasTransfer = True
End Function

Public Function HasRemarks() As Boolean
    If This.Remarks Is Nothing Then Exit Function
    HasRemarks = True
End Function
