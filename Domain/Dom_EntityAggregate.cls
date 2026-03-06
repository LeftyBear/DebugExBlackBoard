VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EntityAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    EntityRows As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal EntityRows As VBA.Collection)
    If EntityRows Is Nothing Then Err.Raise DomErrNothingObject, , "EntityRows is Nothing"
    If EntityRows.Count = 0 Then Err.Raise DomErrEmptyObject, , "EntityRows is empty"
    ValidateRows EntityRows
    Set This.EntityRows = EntityRows
End Sub

Public Property Get Count() As Long
    Count = This.EntityRows.Count
End Property

Private Sub ValidateRows(ByVal EntityRows As VBA.Collection)
    Dim i As Long
    For i = 1 To EntityRows.Count
        If Not TypeOf EntityRows.Item(i) Is Dom_ColumnContext Then
            Err.Raise DomErrInvalidTypeOfObject, , "Invalid row type"
        End If
    Next
End Sub
