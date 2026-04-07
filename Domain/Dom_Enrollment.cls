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
    ID      As Dom_EnrollmentID
    Value   As Dom_EnrollmentValue
End Type

Private This As Member

Public Sub Initialize(ByVal ID As Dom_EnrollmentID, ByVal Value As Dom_EnrollmentValue)
    Set This.ID = ID
    Set This.Value = Value
End Sub

Friend Sub ChangeValueInternal(ByVal NewValue As Dom_EnrollmentValue)
    Set This.Value = NewValue
End Sub

Public Function GetID() As Dom_EnrollmentID
    Set GetID = This.ID
End Function

Public Function GetValue() As Dom_EnrollmentValue
    Set GetValue = This.Value
End Function
