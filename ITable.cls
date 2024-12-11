VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder("Interface")
Option Explicit
Public ParentSheet  As Worksheet
Public Exists       As Boolean
Public ShouldSave   As Boolean
Public Sub Delete()
End Sub
Public Sub Add(ByRef DataBody() As Variant)
End Sub
Public Sub ArrangeHeader(ByVal BasisHeaderFilePath As String)
End Sub
Public Function Target() As ListObject
End Function
Public Function Headers() As Variant()
End Function
Public Function HeaderItemColumns(ByVal ItemName As String) As Long
End Function
Public Function HeaderItemName(ByVal ItemColumns As Long) As String
End Function
Public Function HeaderItemList() As Dictionary
End Function
Public Function Record(ByVal RecordNumber As Long) As Range
End Function
Public Function RecordItem(ByVal RecordNumber As Long, ByVal ItemName As String) As Range
End Function
Public Function RecordItemList(ByVal RecordNumber As Long) As Dictionary
End Function
Public Function ColumnItems(ByVal ItemName As String) As Variant()
End Function

