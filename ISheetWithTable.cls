VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISheetWithTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@IgnoreModule MoveFieldCloserToUsage, ProcedureNotUsed, VariableNotUsed, EncapsulatePublicField
'@Folder("Sheet")
Option Explicit
Public Target As Worksheet
Public Sub TryImport(ByVal SheetFilePath As String, ByVal HeaderFilePath As String)
End Sub
Public Sub TryExport(ByVal SheetFilePath As String)
End Sub
Public Sub AddTable(ByRef DataBody() As Variant)
End Sub
Public Sub DeleteTable()
End Sub
Public Function TargetTable() As ListObject
End Function
Public Function HeaderItemColumns(ByVal ItemName As String) As Long
End Function
Public Function HeaderItemName(ByVal ItemColumns As Long) As String
End Function
Public Function HeaderItemList() As Dictionary
End Function
Public Function DailyRecord(ByVal RecordNumber As Long) As Range
End Function
Public Function DailyRecordItem(ByVal RecordNumber As Long, ByVal ItemName As String) As Range
End Function
Public Function DailyRecordItemList(ByVal RecordNumber As Long) As Dictionary
End Function
Public Function YearlyRecord(ByVal ItemName As String) As Range
End Function
Public Function YearlyRecordItemList(ByVal ItemName As String) As Dictionary
End Function
Public Function HeaderPart() As Range
End Function
Public Function AllRecords() As Range
End Function
Public Function HeaderAndAllRecords() As Range
End Function
Public Function TableExists() As Boolean
End Function

