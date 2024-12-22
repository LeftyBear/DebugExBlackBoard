VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetWithTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sheet")
Option Explicit
Implements ISheetWithTable
Private TargetSheet As Worksheet
Private Table       As ITable
Private Property Set ISheetWithTable_Target(ByVal Target As Worksheet)
    Set TargetSheet = Target
End Property
Private Property Get ISheetWithTable_Target() As Worksheet
    Set ISheetWithTable_Target = TargetSheet
End Property
Private Sub Class_Initialize()
    Set Table = New Table
End Sub
Private Sub ISheetWithTable_TryImport(ByVal SheetFilePath As String, ByVal HeaderFilePath As String)
    ISheetWithTable_DeleteTable
    If Not FSO.FileExists(SheetFilePath) Then Exit Sub
    If Not FSO.FileExists(HeaderFilePath) Then Exit Sub
    Dim DataBody() As Variant
    DataBody = LIB.ReadAllLineAsArray2D(SheetFilePath)
    If LIB.IsNoneArray(DataBody) Then Exit Sub
    If ISheetWithTable_Target.Name Like "予定" Then
        DataBody = ReplaceLineBreakChar(DataBody)
    End If
    Dim LatestHeaderBody() As Variant
    LatestHeaderBody = LIB.ReadAllLineAsArray2D(HeaderFilePath)
    If LIB.IsNoneArray(LatestHeaderBody) Then Exit Sub
    DataBody = ArrangeHeader(LatestHeaderBody, DataBody)
    ISheetWithTable_AddTable DataBody
End Sub
Private Sub ISheetWithTable_TryExport(ByVal SheetFilePath As String)
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Sub
    Dim DataBody() As Variant
    DataBody = ISheetWithTable_AllRecords.Value
    If ISheetWithTable_Target.Name Like "予定" Then
        DataBody = ReplaceLineBreak(DataBody)
    End If
    Dim BUF() As Variant
    BUF = ArrangeDateColumn(DataBody)
    If Not LIB.IsNoneArray(BUF) Then
        DataBody = BUF
        ISheetWithTable_AllRecords.Value = DataBody
    End If
    If UBound(DataBody) < 365 Then
        MsgBox "データ数が不足しているので処理を中断します。" & vbCrLf & _
        "データ破損の可能性があります。システムを再起動してください。", vbCritical, "データ出力エラー"
        Exit Sub
    End If
    Dim HeaderAndAllRecords() As Variant
    HeaderAndAllRecords = ISheetWithTable_HeaderAndAllRecords.Value
    With New ClassBackup
        .WriteFileAfterBackup SheetFilePath, HeaderAndAllRecords
    End With
End Sub
Private Sub ISheetWithTable_AddTable(ByRef DataBody() As Variant)
    Set Table.ParentSheet = ISheetWithTable_Target
    Table.Add DataBody
End Sub
Private Sub ISheetWithTable_DeleteTable()
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Sub
    Table.Delete
End Sub
Private Function ISheetWithTable_TargetTable() As ListObject
    Set Table.ParentSheet = ISheetWithTable_Target
    Set ISheetWithTable_TargetTable = Table.TargetListObject
End Function
Private Function ISheetWithTable_HeaderItemColumns(ByVal ItemName As String) As Long
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    ISheetWithTable_HeaderItemColumns = Table.TargetListObject.ListColumns.Item(ItemName).Index
End Function
Private Function ISheetWithTable_HeaderItemName(ByVal ItemColumns As Long) As String
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    ISheetWithTable_HeaderItemName = Table.TargetListObject.HeaderRowRange.Item(ItemColumns).Value
End Function
Private Function ISheetWithTable_HeaderItemList() As Scripting.IDictionary
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim HeaderBody() As Variant
    HeaderBody = ISheetWithTable_HeaderPart.Value
    Dim C As Long
    For C = LBound(HeaderBody, 2) To UBound(HeaderBody, 2)
        RET.Add HeaderBody(1, C), C
    Next
    Set ISheetWithTable_HeaderItemList = RET
End Function
Private Function ISheetWithTable_DailyRecord(ByVal RecordNumber As Long) As Range
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Set ISheetWithTable_DailyRecord = Table.TargetListObject.ListRows.Item(RecordNumber).Range
End Function
Private Function ISheetWithTable_DailyRecordItem(ByVal RecordNumber As Long, ByVal ItemName As String) As Range
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Dim C As Long
    C = ISheetWithTable_HeaderItemColumns(ItemName)
    Set ISheetWithTable_DailyRecordItem = Table.TargetListObject.ListRows.Item(RecordNumber).Range.Item(C)
End Function
Private Function ISheetWithTable_DailyRecordItemList(ByVal RecordNumber As Long) As Scripting.IDictionary
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim HeaderBody() As Variant
    HeaderBody = ISheetWithTable_HeaderPart.Value
    Dim DailyRecord() As Variant
    DailyRecord = ISheetWithTable_DailyRecord(RecordNumber).Value
    Dim C As Long
    For C = LBound(HeaderBody, 2) To UBound(HeaderBody, 2)
        RET.Add HeaderBody(1, C), DailyRecord(1, C)
    Next
    Set ISheetWithTable_DailyRecordItemList = RET
End Function
Private Function ISheetWithTable_YearlyRecord(ByVal ItemName As String) As Range
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Set ISheetWithTable_YearlyRecord = Table.TargetListObject.ListColumns.Item(ItemName).DataBodyRange
End Function
Private Function ISheetWithTable_YearlyRecordItemList(ByVal ItemName As String) As Scripting.IDictionary
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim ItemBody() As Variant
    ItemBody = ISheetWithTable_YearlyRecord(ItemName).Value
    Dim DateBody() As Variant
    DateBody = ISheetWithTable_YearlyRecord("日付").Value
    Dim i As Long
    For i = LBound(DateBody) To UBound(DateBody)
        RET.Add DateBody(i, 1), ItemBody(i, 1)
    Next
    Set ISheetWithTable_YearlyRecordItemList = RET
End Function
Private Function ISheetWithTable_HeaderPart() As Range
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Set ISheetWithTable_HeaderPart = Table.TargetListObject.HeaderRowRange
End Function
Private Function ISheetWithTable_AllRecords() As Range
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Set ISheetWithTable_AllRecords = Table.TargetListObject.DataBodyRange
End Function
Private Function ISheetWithTable_HeaderAndAllRecords() As Range
    Set Table.ParentSheet = ISheetWithTable_Target
    If Not Table.Exists Then Exit Function
    Set ISheetWithTable_HeaderAndAllRecords = Table.TargetListObject.Range
End Function
Private Function ISheetWithTable_TableExists() As Boolean
    Set Table.ParentSheet = ISheetWithTable_Target
    ISheetWithTable_TableExists = Table.Exists
End Function
Private Function ArrangeHeader(ByRef LatestHeaderBody() As Variant, ByRef DataBody() As Variant) As Variant()
    Dim CurrentHeaderList As Dictionary
    Set CurrentHeaderList = New Dictionary
    Dim C As Long
    For C = 1 To UBound(DataBody, 2)
        CurrentHeaderList.Add DataBody(1, C), C
    Next
    If HeaderMatches(LatestHeaderBody, CurrentHeaderList) Then
        ArrangeHeader = DataBody: Exit Function
    End If
    Dim RET() As Variant
    ReDim RET(1 To UBound(DataBody), 1 To UBound(LatestHeaderBody, 2))
    For C = 1 To UBound(LatestHeaderBody, 2)
        RET(1, C) = LatestHeaderBody(1, C)
        If CurrentHeaderList.Exists(LatestHeaderBody(1, C)) Then
            Dim R As Long
            For R = 2 To UBound(DataBody)
                RET(R, C) = DataBody(R, CurrentHeaderList.Item(LatestHeaderBody(1, C)))
            Next
        End If
    Next
    ArrangeHeader = RET
End Function
Private Function ArrangeDateColumn(ByRef DataBody() As Variant) As Variant()
    Dim RET() As Variant
    RET = DataBody
    Dim C As Long
    C = ISheetWithTable_HeaderItemColumns("日付")
    Dim ArrangedDateList As Dictionary
    Set ArrangedDateList = New Dictionary
    Dim R As Long
    For R = LBound(RET) To UBound(RET)
        Dim KeyDate As Date
        KeyDate = RET(R, C)
        If Not ArrangedDateList.Exists(KeyDate) Then
            Dim Days As Long
            ArrangedDateList.Add KeyDate, Days
        Else
            Dim ExpectDate As Date
            ExpectDate = ArrangedDateList.Item(1) + Days - 1
            RET(R, C) = ExpectDate
        End If
        Days = Days + 1
    Next
    If ExpectDate = 0 Then Exit Function
    ArrangeDateColumn = RET
End Function
Private Function ReplaceLineBreakChar(ByRef DataBody() As Variant) As Variant()
    Dim RET() As Variant
    RET = DataBody
    Dim R As Long
    For R = LBound(RET, 1) To UBound(RET, 1)
        Dim C As Long
        For C = LBound(RET, 2) To UBound(RET, 2)
            If RET(R, C) Like "*" & LINE_BREAK_CHAR & "*" Then
                RET(R, C) = Replace(RET(R, C), LINE_BREAK_CHAR, vbLf)
            End If
        Next
    Next
    ReplaceLineBreakChar = RET
End Function
Private Function ReplaceLineBreak(ByRef DataBody() As Variant) As Variant()
    Dim RET() As Variant
    RET = DataBody
    Dim R As Long
    For R = LBound(RET, 1) To UBound(RET, 1)
        Dim C As Long
        For C = LBound(RET, 2) To UBound(RET, 2)
            RET(R, C) = LIB.LineBreakToVbLf(RET(R, C))
            If InStr(RET(R, C), vbLf) > 0 Then
                RET(R, C) = Replace(RET(R, C), vbLf, LINE_BREAK_CHAR)
            End If
        Next
    Next
    ReplaceLineBreak = RET
End Function
Private Function HeaderMatches(ByRef LatestHeaderBody() As Variant, ByVal CurrentHeaderList As Dictionary) As Boolean
    Dim C As Long
    For C = 1 To UBound(LatestHeaderBody, 2)
        If Not CurrentHeaderList.Exists(LatestHeaderBody(1, C)) Then Exit Function
        If C <> CurrentHeaderList.Item(LatestHeaderBody(1, C)) Then Exit Function
    Next
    HeaderMatches = True
End Function

