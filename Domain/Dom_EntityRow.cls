VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EntityRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    EntityColumns As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal EntityColumns As VBA.Collection)
    If EntityColumns Is Nothing Then Err.Raise DomErrNothingObject, , "EntityColumns is Nothing"
    If EntityColumns.Count = 0 Then Err.Raise DomErrEmptyObject, , "EntityColumns is empty"
    ValidateColumns EntityColumns
    Set This.EntityColumns = EntityColumns
End Sub

Public Property Get EntityColumns() As VBA.Collection
    Set EntityColumns = This.EntityColumns
End Property

Public Property Get Count() As Long
    Count = This.EntityColumns.Count
End Property

Public Function EntityColumn(ByVal Index As Long) As Dom_IEntityColumn
    Set EntityColumn = This.EntityColumns.Item(Index)
End Function

Private Sub ValidateColumns(ByVal Source As VBA.Collection)
    Dim i As Long
    For i = 1 To Source.Count
        If Not TypeOf Source.Item(i) Is Dom_IEntityColumn Then Err.Raise DomErrInvalidTypeOfObject, , "Invalid column type"
    Next
End Sub
