VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.DTOCollection")
Option Explicit
Private Type Member
    Rows As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Rows = New VBA.Collection
End Sub

Public Sub Add(ByVal Row As Inf_SubjectRow)
    This.Rows.Add Row
End Sub

Public Sub Remove(ByVal INDEX As Long)
    This.Rows.Remove INDEX
End Sub

Public Function Count() As Long
    Count = This.Rows.Count
End Function

Public Function Item(ByVal INDEX As Long) As Inf_SubjectRow
    Set Item = This.Rows.Item(INDEX)
End Function
