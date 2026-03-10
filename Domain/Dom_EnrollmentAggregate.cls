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
    Records As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal Records As VBA.Collection)
    If Records Is Nothing Then Err.Raise DomErrNothingObject, , "Records is Nothing"
    If Records.Count = 0 Then Err.Raise DomErrEmptyObject, , "Records is empty"
    ValidateRecord Records
    Set This.Records = Records
End Sub

Public Property Get Count() As Long
    Count = This.Records.Count
End Property

Public Function GetRecord(ByVal Index As Long) As Dom_EnrollmentRecord
    Set GetRecord = This.Records.Item(Index)
End Function

Private Sub ValidateRecord(ByVal Records As VBA.Collection)
    Dim i As Long
    For i = 1 To Records.Count
        If Not TypeOf Records.Item(i) Is Dom_EnrollmentRecord Then
            Err.Raise DomErrInvalidTypeOfObject, , "Invalid Dom_EnrollmentRecord type"
        End If
    Next
End Sub
