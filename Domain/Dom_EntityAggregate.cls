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
    ColumnContextContainer As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal ColumnContextContainer As VBA.Collection)
    If ColumnContextContainer Is Nothing Then Err.Raise DomErrNothingObject, , "ColumnContextContainer is Nothing"
    If ColumnContextContainer.Count = 0 Then Err.Raise DomErrEmptyObject, , "ColumnContextContainer is empty"
    ValidateContainer ColumnContextContainer
    Set This.ColumnContextContainer = ColumnContextContainer
End Sub

Public Property Get Count() As Long
    Count = This.ColumnContextContainer.Count
End Property

Private Sub ValidateContainer(ByVal ColumnContextContainer As VBA.Collection)
    Dim i As Long
    For i = 1 To ColumnContextContainer.Count
        If Not TypeOf ColumnContextContainer.Item(i) Is Dom_ColumnContext Then
            Err.Raise DomErrInvalidTypeOfObject, , "Invalid row type"
        End If
    Next
End Sub
