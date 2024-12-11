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
Private Exists      As Boolean
Private ShouldSave  As Boolean
Private Property Set ITable_ParentSheet(ByVal TargetSheet As Worksheet)
    Set ParentSheet = TargetSheet
End Property
Private Property Get ITable_ParentSheet() As Worksheet
    Set ITable_ParentSheet = ParentSheet
End Property
Private Property Let ITable_Exists(ByVal State As Boolean)
    Exists = State
End Property
Private Property Get ITable_Exists() As Boolean
    If ITable_Target Is Nothing Then Exit Property
    ITable_Exists = True
End Property
Private Property Let ITable_ShouldSave(ByVal State As Boolean)
    ShouldSave = State
End Property
Private Property Get ITable_ShouldSave() As Boolean
    ITable_ShouldSave = ShouldSave
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
        ITable_Exists = True
    End With
End Sub
Private Sub ITable_Delete()
    If Not ITable_Exists Then Exit Sub
    ITable_Target.Delete
    ITable_Exists = False
End Sub
Private Sub ITable_ArrangeHeader(ByVal BasisHeaderFilePath As String)
    If Not ITable_Exists Then Exit Sub
    Dim BasisHeaders() As Variant
    BasisHeaders = LIB.ReadAllLineAsArray2D(BasisHeaderFilePath)
    If LIB.NotIsArray(BasisHeaders) Then Exit Sub
    DeleteHeaderItem BasisHeaders
    AddHeaderItem BasisHeaders
    SortHeaderItem BasisHeaders
End Sub
Private Sub DeleteHeaderItem(ByRef BasisHeaders() As Variant)
    With ITable_Target
        Dim Headers() As Variant
        Headers = .HeaderRowRange.Value
        Dim C As Long
        For C = LBound(Headers, 2) To UBound(Headers, 2)
            Dim BasisC As Long
            For BasisC = LBound(BasisHeaders, 2) To UBound(BasisHeaders, 2)
                If Headers(1, C) Like BasisHeaders(1, BasisC) Then
                    Dim HasFound As Boolean
                    HasFound = True
                    Exit For
                End If
            Next
            If Not HasFound Then
                .ListColumns.Item(Headers(1, C)).Range.Delete
                ITable_ShouldSave = True
            End If
            HasFound = False
        Next
    End With
End Sub
Private Sub AddHeaderItem(ByRef BasisHeaders() As Variant)
    With ITable_Target
        Dim BasisC As Long
        For BasisC = LBound(BasisHeaders, 2) To UBound(BasisHeaders, 2)
            If WorksheetFunction.CountIf(.HeaderRowRange, BasisHeaders(1, BasisC)) = 0 Then
                .ListColumns.Add(BasisC).Range.Item(1) = BasisHeaders(1, BasisC)
                ITable_ShouldSave = True
            End If
        Next
    End With
End Sub
Private Sub SortHeaderItem(ByRef BasisHeaders() As Variant)
    With ITable_Target
        Dim Headers() As Variant
        Headers = .HeaderRowRange.Value
        If UBound(Headers, 2) <> UBound(BasisHeaders, 2) Then
            MsgBox "ヘッダーの要素数が一致していないため、並べ替えられません。", vbExclamation, _
                "コーディングエラー": Exit Sub
        End If
        Dim DefaultIndex As Long
        For DefaultIndex = LBound(BasisHeaders, 2) To UBound(BasisHeaders, 2)
            If Not BasisHeaders(1, DefaultIndex) Like Headers(1, DefaultIndex) Then
                Dim CurrentIndex As Long
                For CurrentIndex = DefaultIndex + 1 To UBound(BasisHeaders, 2)
                    If BasisHeaders(1, DefaultIndex) Like Headers(1, CurrentIndex) Then
                        .ListColumns.Item(CurrentIndex).Range.Cut
                        .ListColumns.Item(DefaultIndex).Range.Insert Shift:=xlToRight
                        Application.CutCopyMode = False
                        Dim BUF As Variant
                        BUF = Headers(1, DefaultIndex)
                        Headers(1, DefaultIndex) = Headers(1, CurrentIndex)
                        Headers(1, CurrentIndex) = BUF
                        ITable_ShouldSave = True
                        Exit For
                    End If
                Next
            End If
        Next
    End With
End Sub
Private Function ITable_Target() As ListObject
    With ITable_ParentSheet.ListObjects
        If .Count = 0 Then ITable_Exists = False: Exit Function
        Set ITable_Target = .Item(1)
        ITable_Exists = True
    End With
End Function
Private Function ITable_Headers() As Variant()
    If Not ITable_Exists Then Exit Function
    ITable_Headers = ITable_Target.HeaderRowRange.Value
End Function
Private Function ITable_HeaderItemColumns(ByVal ItemName As String) As Long
    If Not ITable_Exists Then Exit Function
    ITable_HeaderItemColumns = ITable_Target.ListColumns.Item(ItemName).Index
End Function
Private Function ITable_HeaderItemName(ByVal ItemColumns As Long) As String
    If Not ITable_Exists Then Exit Function
    ITable_HeaderItemName = ITable_Target.HeaderRowRange.Item(ItemColumns).Value
End Function
Private Function ITable_HeaderItemList() As Dictionary
    If Not ITable_Exists Then Exit Function
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim Headers() As Variant
    Headers = ITable_Headers
    Dim C As Long
    For C = LBound(Headers, 2) To UBound(Headers, 2)
        RET.Add Headers(1, C), C
    Next
    Set ITable_HeaderItemList = RET
End Function
Private Function ITable_Record(ByVal RecordNumber As Long) As Range
    If Not ITable_Exists Then Exit Function
    Set ITable_Record = ITable_Target.ListRows.Item(RecordNumber).Range
End Function
Private Function ITable_RecordItem(ByVal RecordNumber As Long, ByVal ItemName As String) As Range
    If Not ITable_Exists Then Exit Function
    Dim C As Long
    C = ITable_HeaderItemColumns(ItemName)
    Set ITable_RecordItem = ITable_Record(RecordNumber).Item(C)
End Function
Private Function ITable_RecordItemList(ByVal RecordNumber As Long) As Dictionary
    If Not ITable_Exists Then Exit Function
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim Headers() As Variant
    Headers = ITable_Headers
    Dim Records() As Variant
    Records = ITable_Record(RecordNumber).Value
    Dim C As Long
    For C = LBound(Headers, 2) To UBound(Headers, 2)
        RET.Add Headers(1, C), Records(1, C)
    Next
    Set ITable_RecordItemList = RET
End Function
Private Function ITable_ColumnItems(ByVal ItemName As String) As Variant()
    If Not ITable_Exists Then Exit Function
    ITable_ColumnItems = ITable_Target.ListColumns.Item(ItemName).DataBodyRange.Value
End Function

