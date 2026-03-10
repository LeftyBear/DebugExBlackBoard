VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ConfigColumnFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit
Implements Dom_IEntityColumnFactory

Private Function Dom_IEntityColumnFactory_Create(ByVal ColumnDTO As App_ColumnDTO) As Dom_IValueObject
    Select Case ColumnDTO.HeaderName
    Case "普通学年数", "普通学級数", "特支学年数", "特支学級数", "時間割数"
        Dim LongValueColumn As Dom_LongValueColumn
        Set LongValueColumn = New Dom_LongValueColumn
        Set Dom_IEntityColumnFactory_Create = LongValueColumn.Create(ColumnDTO.HeaderName, CLng(ColumnDTO.RawValue))
    Case "集計名", "対象学年"
        Dim ScalarColumn As Dom_StringValueColumn
        Set ScalarColumn = New Dom_StringValueColumn
        Set Dom_IEntityColumnFactory_Create = ScalarColumn.Create(ColumnDTO.HeaderName, CLng(ColumnDTO.RawValue))
    End Select
End Function
