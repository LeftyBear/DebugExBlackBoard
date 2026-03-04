VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ConfigRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    CompositeColumns As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal CompositeColumns As VBA.Collection)
    If CompositeColumns Is Nothing Then Err.Raise DomErrNothingObject, , "CompositeColumns is Nothing"
    If CompositeColumns.Count = 0 Then Err.Raise DomErrEmptyObject, , "CompositeColumns is empty"
    ValidateColumns CompositeColumns
    Set This.CompositeColumns = CompositeColumns
End Sub

Public Property Get CompositeColumns() As VBA.Collection
    Set CompositeColumns = This.CompositeColumns
End Property

Public Property Get Count() As Long
    Count = This.CompositeColumns.Count
End Property

Public Function CompositeColumn(ByVal Index As Long) As Dom_IEntityColumn
    Set CompositeColumn = This.CompositeColumns.Item(Index)
End Function

Private Sub ValidateColumns(ByVal Source As VBA.Collection)
    Dim i As Long
    For i = 1 To Source.Count
        If Not TypeOf Source.Item(i) Is Dom_IEntityColumn Then Err.Raise DomErrInvalidTypeOfObject, , "Invalid column type"
    Next
End Sub
