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
        LongValueColumn.Initialize ColumnDTO.HeaderName, CLng(ColumnDTO.RawValue)
        Set Dom_IEntityColumnFactory_Create = LongValueColumn
    Case "集計名", "対象学年"
        Dim ScalarColumn As Dom_StringValueColumn
        Set ScalarColumn = New Dom_StringValueColumn
        ScalarColumn.Initialize ColumnDTO.HeaderName, ColumnDTO.RawValue
        Set Dom_IEntityColumnFactory_Create = ScalarColumn
    End Select
End Function
