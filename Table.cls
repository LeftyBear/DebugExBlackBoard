VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Table")
Option Explicit
Implements ITable
Private ParentSheet As Worksheet
Private Property Set ITable_ParentSheet(ByVal TargetSheet As Worksheet)
    Set ParentSheet = TargetSheet
End Property
Private Property Get ITable_ParentSheet() As Worksheet
    Set ITable_ParentSheet = ParentSheet
End Property
Private Sub ITable_Add(ByRef DataBody() As Variant)
    With ITable_ParentSheet
        Dim TargetRange As Range
        Set TargetRange = .Range("A1").Resize(UBound(DataBody, 1), UBound(DataBody, 2))
        TargetRange.Value = DataBody
        With .ListObjects.Add(xlSrcRange, TargetRange, , xlYes)
            .TableStyle = vbNullString
            .ShowAutoFilter = False
        End With
    End With
End Sub
Private Sub ITable_Delete()
    ITable_TargetListObject.Delete
End Sub
Private Function ITable_TargetListObject() As ListObject
    With ITable_ParentSheet.ListObjects
        If .Count = 0 Then Exit Function
        Set ITable_TargetListObject = .Item(1)
    End With
End Function
Private Function ITable_Exists() As Boolean
    If ITable_TargetListObject Is Nothing Then Exit Function
    ITable_Exists = True
End Function

