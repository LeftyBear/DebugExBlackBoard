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
    Rows As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal Rows As VBA.Collection)
    If Rows Is Nothing Then Err.Raise errNothingRows, , "Rows is Nothing"
    If Rows.Count = 0 Then Err.Raise errEmptyRows, , "Rows is empty"
    ValidateRows Rows
    Set This.Rows = Rows
End Sub

Public Property Get Rows() As VBA.Collection
    Set Rows = This.Rows
End Property

Public Property Get Count() As Long
    Count = This.Rows.Count
End Property

Private Sub ValidateRows(ByVal Rows As VBA.Collection)
    Dim i As Long
    For i = 1 To Rows.Count
        If Not TypeOf Rows.Item(i) Is Dom_EnrollmentRow Then
            Err.Raise errInvalidRowType, , "Invalid row type"
        End If
    Next
End Sub

