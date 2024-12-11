Attribute VB_Name = "LIB"
'@IgnoreModule EncapsulatePublicField, ProcedureCanBeWrittenAsFunction, UseMeaningfulName, WriteOnlyProperty, ProcedureNotUsed, MultipleDeclarations, IIfSideEffect
'@Folder("Library")
Option Explicit
Option Private Module
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If
Public FSO As New FileSystemObject
Public Enum SORT_ORDER
    xlAscending = 1
    xlDescending
End Enum
Public Enum DATE_FORMAT
    Western = 1
    Japanese
    Japanese強制
End Enum
Public Enum RGB_COLOR
    Black = 0
    White = 16777215
    Red = 255
    Yellow = 65535
    Blue = 16711680
    LightGreen = 10156544
    DarkGray = 11119017
End Enum
Public Const AS_FOLDER          As String = "\"
Public Const CANNMA             As String = ","
Public Const SEMICOLON          As String = ";"
Public Const DOUBLE_QUOTATION   As String = """"
Public Const WIDE_SPACE         As String = "　"
Public Const HALF_SPACE         As String = " "
Public Const LINE_BREAK_CHAR    As String = "<br>"
Public Sub AddItemOfListBox(ByVal ListBox As MSForms.ListBox, ByVal Target As String)
    With ListBox
        Dim i As Long
        For i = 0 To .ListCount - 1
            .Selected(i) = False
            If .List(i) = Target Then
                .Selected(i) = True
                Dim ExistsList As Boolean
                ExistsList = True
            End If
        Next
        If ExistsList Then Exit Sub
        .AddItem Target
        .Selected(.ListCount - 1) = True
        .ListIndex = .ListCount - 1
    End With
End Sub
Public Property Let Application画面更新(ByVal State As Boolean)
    With Application
        .ScreenUpdating = State
        .Cursor = IIf(State, xlDefault, xlWait)
        .Calculation = IIf(State, xlCalculationAutomatic, xlCalculationManual)
    End With
End Property
Public Function CellDiff(ByVal Sheet As Worksheet, ByVal OutsideRange As Range, ByVal InsideRange As Range) As String
    Dim RowDiff As Long, ColumnDiff As Long
    RowDiff = InsideRange.Row - OutsideRange.Item(1).Row + 1
    ColumnDiff = InsideRange.Column - OutsideRange.Item(1).Column + 1
    CellDiff = Sheet.Cells.Item(RowDiff, ColumnDiff).Address(False, False)
End Function
Public Function CollectionItemsToArray(ByVal CollectionObject As Collection) As Variant()
    Dim RET() As Variant
    ReDim RET(CollectionObject.Count - 1)
    Dim i As Long
    For i = 0 To CollectionObject.Count - 1
        RET(i) = CollectionObject.Item(i + 1)
    Next
    CollectionItemsToArray = RET
End Function
Public Sub CreateFolder(ByVal FolderPath As String, Optional ByVal HiddenMode As Boolean)
    Dim ParentFolderPath As String
    ParentFolderPath = FSO.GetParentFolderName(FolderPath)
    If Not FSO.FolderExists(ParentFolderPath) Then CreateFolder ParentFolderPath, HiddenMode
    If Not FSO.FolderExists(FolderPath) Then
        FSO.CreateFolder FolderPath
        If HiddenMode Then FSO.GetFolder(FolderPath).Attributes = Hidden
    End If
End Sub
Public Function Date今週初日(ByVal TargetDate As Date, Optional ByVal BeginWeekDay As Long = vbMonday) As Date
    Dim BUF As Date
    BUF = TargetDate
    If Weekday(BUF) = BeginWeekDay And BeginWeekDay = vbSaturday Then BUF = TargetDate - 7
    Date今週初日 = BUF - Weekday(BUF, BeginWeekDay) + 1
End Function
Public Function Date今週末日(ByVal TargetDate As Date, Optional ByVal EndWeekDay As Long = vbSunday) As Date
    Date今週末日 = TargetDate - Weekday(TargetDate, EndWeekDay) + 1 + 7
End Function
Public Function Date今月1日(ByVal TargetDate As Date) As Date
    Date今月1日 = WorksheetFunction.EoMonth(TargetDate, -1) + 1
End Function
Public Function Date今月末日(ByVal TargetDate As Date) As Date
    Date今月末日 = WorksheetFunction.EoMonth(TargetDate, 0)
End Function
Public Function Date4月1日(ByVal TargetDate As Date) As Date
    Select Case Month(TargetDate)
        Case 4 To 12
            Dim RET As Date
            RET = DateSerial(Year(TargetDate), 4, 1)
        Case 1 To 3
            RET = DateSerial(Year(TargetDate) - 1, 4, 1)
    End Select
    Date4月1日 = RET
End Function
Public Function Date3月31日(ByVal TargetDate As Date) As Date
    Select Case Month(TargetDate)
        Case 4 To 12
            Dim RET As Date
            RET = DateSerial(Year(TargetDate) + 1, 3, 31)
        Case 1 To 3
            RET = DateSerial(Year(TargetDate), 3, 31)
    End Select
    Date3月31日 = RET
End Function
Public Function DateDiff4月1日(ByVal Target As Date, Optional ByVal IncludesTarget As Boolean = True) As Long
    If Target = 0 Then Exit Function
    Dim AddDay As Long
    If IncludesTarget Then AddDay = 1
    DateDiff4月1日 = Target - LIB.Date4月1日(Target) + AddDay
End Function
Public Function EqualsTextString(ByVal OneText As String, ByVal AnotherText As String) As Boolean
    If Not OneText Like AnotherText Then Exit Function
    EqualsTextString = True
End Function
Public Function ErasureFromLeft(ByVal Text As String, ByVal CharLength As Long) As String
    If CharLength < 0 Then Exit Function
    ErasureFromLeft = Mid$(Text, Int(CharLength) + 1)
End Function
Public Function ErasureFromRight(ByVal Text As String, ByVal CharLength As Long) As String
    If CharLength < 0 Then Exit Function
    ErasureFromRight = Left$(Text, Len(Text) - Int(CharLength))
End Function
Public Sub ExecuteShellCommand(ByVal Commandlet As String, Optional ByRef IsFailed As Boolean = False)
    On Error GoTo ThrowError
    Dim ShellObject As IWshRuntimeLibrary.WshShell
    Set ShellObject = New IWshRuntimeLibrary.WshShell
    Dim ShellCommand As IWshRuntimeLibrary.WshExec
    Set ShellCommand = ShellObject.Exec(Commandlet)
    If ShellCommand.Status = WshFailed Then
        IsFailed = True
        Exit Sub
    Else
        Dim StartTime As Single
        StartTime = Timer
        Const LimitTime As Single = 10#
        Do While ShellCommand.Status = WshRunning
            DoEvents
            If Timer > StartTime + LimitTime Then
                IsFailed = True
                Exit Do
            End If
        Loop
    End If
Exit Sub
ThrowError:
    IsFailed = True
End Sub
Public Function FetchURLFile(ByVal URL As String, ByVal FilePath As String) As Long
    FetchURLFile = URLDownloadToFile(0, URL, FilePath, 0, 0)
End Function
Public Sub FileErrorRise(ByVal FilePath As String)
    Err.Raise 53, , FSO.GetFileName(FilePath) & "が見つかりません。"
End Sub
Public Function FileDialogFilePicker(Optional ByVal MultiSelect As Boolean) As Variant
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "選択可能ファイル", "*.xls;*.xlsx;*.csv", 1
        .AllowMultiSelect = MultiSelect
        .Title = "ファイルの選択"
        If .Show = False Then Exit Function
        Dim i As Long, FileName() As String
        For i = 1 To .SelectedItems.Count
            ReDim Preserve FileName(i - 1)
            FileName(i - 1) = .SelectedItems.Item(i)
        Next
    End With
    FileDialogFilePicker = FileName
End Function
Public Sub FolderErrorRise(ByVal FolderPath As String)
    Err.Raise 76, , FSO.GetFolder(FolderPath).Name & "が見つかりません。"
End Sub
Public Function Format日付表記(ByVal TargetDate As Date, ByVal Choice As DATE_FORMAT) As String
    Dim RET As String
    Select Case Choice
        Case DATE_FORMAT.Western
            RET = "Long Date"
        Case DATE_FORMAT.Japanese
            RET = "ggge年m月d日(aaa)"
        Case DATE_FORMAT.Japanese強制
            RET = Format強制令和表記(TargetDate)
    End Select
    Format日付表記 = RET
End Function
Private Function Format強制令和表記(ByVal TargetDate As Date) As String
    Dim RET As String
    Select Case TargetDate
        Case #5/1/2019# To #12/31/2019#
            RET = "令和元年" & Format$(TargetDate, "m月d日(aaa)")
        Case Is >= #1/1/2020#
            RET = "令和" & Format$(TargetDate, "yyyy") - 2018 & "年" & Format$(TargetDate, "m月d日(aaa)")
        Case Else
            RET = Format$(TargetDate, "ggge年m月d日(aaa)")
    End Select
    Format強制令和表記 = RET
End Function
Public Function GetBaseNamesInFolder(ByVal FolderPath As String, Optional ByVal KeyWord As String) As Variant()
    If FSO.FolderExists(FolderPath) Then Exit Function
    Dim RET() As Variant
    Dim Target As File
    For Each Target In FSO.GetFolder(FolderPath).Files
        If Not KeyWord Like vbNullString Then
            If FSO.GetBaseName(Target.Name) Like "*" & KeyWord & "*" Then
                Dim i As Long
                ReDim Preserve RET(i)
                RET(i) = FSO.GetBaseName(Target.Name)
                i = i + 1
            End If
        Else
            ReDim Preserve RET(i)
            RET(i) = FSO.GetBaseName(Target.Name)
            i = i + 1
        End If
    Next
    GetBaseNamesInFolder = RET
End Function
Public Function HasIllegalChar(ByVal Text As String, ByVal Letter As String) As Boolean
    If Text Like vbNullString Or Letter Like vbNullString Then Exit Function
    If InStr(Text, Letter) = 0 Then Exit Function
    HasIllegalChar = True
    MsgBox "禁則文字「" & Letter & "」は使用できません。", vbExclamation, "文字入力エラー"
End Function
Public Function HasWideSpaceOfMid(ByVal Text As String) As Boolean
    Dim StringsUnit() As String
    StringsUnit = Split(Text, WIDE_SPACE)
    If UBound(StringsUnit) <> 1 Then Exit Function
    Dim i As Long
    For i = LBound(StringsUnit) To UBound(StringsUnit)
        If Not StringsUnit(i) Like vbNullString Then
            Dim Units As Long
            Units = Units + 1
        End If
    Next
    If Units = 2 Then HasWideSpaceOfMid = True
End Function
Public Function HttpResponseHeader(ByVal URL As String, ByVal HeaderTag As String) As String
    Dim HttpObject As XMLHTTP60
    Set HttpObject = New XMLHTTP60
'    Dim HttpObject As Object
'    Set HttpObject = CreateObject("MSXML2.HTTP")
    On Error Resume Next
    HttpObject.Open "HEAD", URL, False
    HttpObject.send
    Do While HttpObject.readyState < 4
        DoEvents
    Loop
    Dim Headers As String
    Headers = HttpObject.getAllResponseHeaders()
    On Error GoTo 0
    If Not Headers Like "*" & HeaderTag & "*" Then Exit Function
    HttpResponseHeader = HttpObject.getResponseHeader(HeaderTag)
End Function
Public Function InString(ByVal Text As String, ParamArray Letters()) As Boolean
    If Text Like vbNullString Then Exit Function
    If UBound(Letters) = -1 Then Exit Function
    Dim i As Long
    For i = LBound(Letters) To UBound(Letters)
        If InStr(Text, Letters(i)) Then InString = True
    Next
End Function
Public Function IsHalfNumber(ByVal Text As String) As Boolean
    If Text Like vbNullString Then Exit Function
    With New RegExp
        .Pattern = "^[0-9]+$"
        .Global = True
        If Not .test(Text) Then Exit Function
    End With
    IsHalfNumber = True
End Function
Public Function IsNarrow(ByVal Text As String) As Boolean
    If Not Text Like StrConv(Text, vbNarrow) Then Exit Function
    IsNarrow = True
End Function
Public Function IsOpeningBook(ByVal BookName As String) As Boolean
    If BookName Like vbNullString Then Exit Function
    Dim Book As Workbook
    For Each Book In Workbooks
        If Book.Name = BookName Then
            IsOpeningBook = True
            Exit Function
        End If
    Next
End Function
Public Function JoinShellCommandForCompress(ByVal TargetPath As String, ByVal DestinationPath As String) As String
    Dim Notes(9) As String
    Notes(0) = "powers" & "hell"
    Notes(1) = "-NoLogo"
    Notes(2) = "-ExecutionPolicy RemoteSigned"
    Notes(3) = "-Command"
    Notes(4) = "Compress-Archive"
    Notes(5) = "-Path"
    Notes(6) = TargetPath
    Notes(7) = "-DestinationPath"
    Notes(8) = DestinationPath
    Notes(9) = "-Force"
    JoinShellCommandForCompress = Join(Notes, HALF_SPACE)
End Function
Public Function JoinShellCommandForExpand(ByVal TargetPath As String, ByVal DestinationPath As String) As String
    Dim Notes(9) As String
    Notes(0) = "powers" & "hell"
    Notes(1) = "-NoLogo"
    Notes(2) = "-ExecutionPolicy RemoteSigned"
    Notes(3) = "-Command"
    Notes(4) = "Expand-Archive"
    Notes(5) = "-Path"
    Notes(6) = TargetPath
    Notes(7) = "-DestinationPath"
    Notes(8) = DestinationPath
    Notes(9) = "-Force"
    JoinShellCommandForExpand = Join(Notes, HALF_SPACE)
End Function
Public Function LineBreakCharToVbLf(ByVal Text As String) As String
    Dim RET As String
    If InStr(Text, LINE_BREAK_CHAR) > 0 Then
        RET = Replace(Text, LINE_BREAK_CHAR, vbLf)
    Else
        RET = Text
    End If
    LineBreakCharToVbLf = RET
End Function
Public Function LineBreakToVbLf(ByVal Text As String) As String
    Dim RET As String
    RET = Text
    If InStr(RET, vbCrLf) > 0 Then RET = Replace(RET, vbCrLf, vbLf)
    If InStr(RET, vbCr) > 0 Then RET = Replace(RET, vbCr, vbLf)
    LineBreakToVbLf = RET
End Function
Public Function ListViewItemsToArray(ByVal ListViewObject As MSComctlLib.ListView) As Variant()
    If ListViewObject.ListItems.Count = 0 Then Exit Function
    Dim MaxR As Long, MaxC As Long
    MaxR = ListViewObject.ListItems.Count
    MaxC = ListViewObject.ColumnHeaders.Count
    Dim RET() As Variant
    ReDim RET(1 To MaxR, 1 To MaxC)
    Dim R As Long, C As Long
    For R = 1 To MaxR
        RET(R, 1) = ListViewObject.ListItems.Item(R).Text
        For C = 2 To MaxC
            RET(R, C) = ListViewObject.ListItems.Item(R).SubItems(C - 1)
        Next
    Next
    ListViewItemsToArray = RET
End Function
Public Function LookForArrayIndex(ByRef TargetArray() As Variant, ByVal Target As Variant) As Long
    Dim MinIndex As Long, MaxIndex As Long
    MinIndex = LBound(TargetArray)
    MaxIndex = UBound(TargetArray)
    Do While MinIndex <= MaxIndex
        Dim MidIndex As Long
        MidIndex = Int((MaxIndex + MinIndex) / 2)
        If TargetArray(MidIndex) = Target Then
            LookForArrayIndex = MidIndex
            Exit Function
        ElseIf TargetArray(MidIndex) < Target Then
            MinIndex = MidIndex + 1
        Else
            MaxIndex = MidIndex - 1
        End If
    Loop
End Function
Public Function MergeDictionary(ByVal BaseDictionary As Dictionary, ByVal AddDictionary As Dictionary) As Dictionary
    Dim RET As Dictionary
    Set RET = BaseDictionary
    Dim AddKey As Variant
    For Each AddKey In AddDictionary
        If Not RET.Exists(AddKey) Then
            RET.Add AddKey, AddDictionary.Item(AddKey)
        End If
    Next
    Set MergeDictionary = RET
End Function
Public Function MergeVArray2D(ByRef BaseArray2D() As Variant, ByRef AddArray2D() As Variant) As Variant()
    Dim MaxR As Long, MaxC As Long
    MaxR = UBound(BaseArray2D, 1) + UBound(AddArray2D, 1)
    MaxC = Application.WorksheetFunction.Max(UBound(BaseArray2D, 2), UBound(AddArray2D, 2))
    Dim RET() As Variant
    ReDim RET(1 To MaxR, 1 To MaxC)
    Dim R As Long, C As Long
    For R = 1 To MaxR
        If R <= UBound(BaseArray2D, 1) Then
            For C = 1 To MaxC
                If C <= UBound(BaseArray2D, 2) Then
                    RET(R, C) = BaseArray2D(R, C)
                Else
                    RET(R, C) = Empty
                End If
            Next
        Else
            For C = 1 To MaxC
                If C <= UBound(AddArray2D, 2) Then
                    RET(R, C) = AddArray2D(R - UBound(BaseArray2D, 1), C)
                Else
                    RET(R, C) = Empty
                End If
            Next
        End If
    Next
    MergeVArray2D = RET
End Function
Public Function MergeHArray2D(ByRef BaseArray2D() As Variant, ByRef AddArray2D() As Variant) As Variant()
    Dim MaxR As Long, MaxC As Long
    MaxR = Application.WorksheetFunction.Max(UBound(BaseArray2D, 1), UBound(AddArray2D, 1))
    MaxC = UBound(BaseArray2D, 2) + UBound(AddArray2D, 2)
    Dim RET() As Variant
    ReDim RET(1 To MaxR, 1 To MaxC)
    Dim R As Long, C As Long
    For C = 1 To MaxC
        If C <= UBound(BaseArray2D, 2) Then
            For R = 1 To MaxR
                If R <= UBound(BaseArray2D, 1) Then
                    RET(R, C) = BaseArray2D(R, C)
                Else
                    RET(R, C) = Empty
                End If
            Next
        Else
            For R = 1 To MaxR
                If R <= UBound(AddArray2D, 1) Then
                    RET(R, C) = AddArray2D(R, C - UBound(BaseArray2D, 2))
                Else
                    RET(R, C) = Empty
                End If
            Next
        End If
    Next
    MergeHArray2D = RET
End Function
Public Function MergeSort(ByRef Array1D() As Variant, Optional ByVal Order As SORT_ORDER = xlAscending) As Variant()
    Dim RET() As Variant
    Dim PreviousArray() As Variant, NextArray() As Variant
    If UBound(Array1D) > 0 Then
        Dim Middle As Long
        Middle = Int((UBound(Array1D) + 1) / 2) - 1
        Dim i As Long
        i = 0
        Dim IndexP As Long
        For IndexP = 0 To Middle
            ReDim Preserve PreviousArray(i)
            PreviousArray(i) = Array1D(IndexP)
            i = i + 1
        Next
        i = 0
        Dim IndexN As Long
        For IndexN = Middle + 1 To UBound(Array1D)
            ReDim Preserve NextArray(i)
            NextArray(i) = Array1D(IndexN)
            i = i + 1
        Next
        PreviousArray = MergeSort(PreviousArray, Order)
        NextArray = MergeSort(NextArray, Order)
        RET = MergeArrayForMergeSort(PreviousArray, NextArray, Order)
    Else
        RET = Array1D
    End If
    MergeSort = RET
End Function
Private Function MergeArrayForMergeSort(ByRef PreviousArray() As Variant, ByRef NextArray() As Variant, Optional ByVal Order As SORT_ORDER = xlAscending) As Variant()
    Dim RET() As Variant
    If Order = xlAscending Then
        Dim IndexP As Long, IndexN As Long
        Do While (IndexP <= UBound(PreviousArray)) And (IndexN <= UBound(NextArray))
            If PreviousArray(IndexP) >= NextArray(IndexN) Then
                Dim BUF As Variant
                BUF = NextArray(IndexN)
                IndexN = IndexN + 1
            Else
                BUF = PreviousArray(IndexP)
                IndexP = IndexP + 1
            End If
            Dim i As Long
            ReDim Preserve RET(i)
            RET(i) = BUF
            i = i + 1
        Loop
    Else
        Do While (IndexP <= UBound(PreviousArray)) And (IndexN <= UBound(NextArray))
            If PreviousArray(IndexP) <= NextArray(IndexN) Then
                BUF = NextArray(IndexN)
                IndexN = IndexN + 1
            Else
                BUF = PreviousArray(IndexP)
                IndexP = IndexP + 1
            End If
            ReDim Preserve RET(i)
            RET(i) = BUF
            i = i + 1
        Loop
    End If
    If IndexP <> UBound(PreviousArray) + 1 Then
        Dim j As Long
        For j = IndexP To UBound(PreviousArray)
            ReDim Preserve RET(i)
            RET(i) = PreviousArray(j)
            i = i + 1
        Next
    End If
    If IndexN <> UBound(NextArray) + 1 Then
        For j = IndexN To UBound(NextArray)
            ReDim Preserve RET(i)
            RET(i) = NextArray(j)
            i = i + 1
        Next
    End If
    MergeArrayForMergeSort = RET
End Function
Public Function MidFromLeftToBeforeTargetChar(ByVal Text As String, ByVal Letter As String) As String
    If Text Like vbNullString Then Exit Function
    MidFromLeftToBeforeTargetChar = Left$(Text, InStr(Text, Letter) - 1)
End Function
Public Function MidFromTargetCharToTargetChar(ByVal Text As String, ByVal FirstLetter As String, ByVal LastLetter As String, _
Optional ByVal ContainsFirstLetter As Boolean, Optional ByVal ContainsLastLetter As Boolean) As String
    Dim BeginingPosition As Long, EndPosition As Long
    BeginingPosition = InStr(Text, FirstLetter)
    EndPosition = InStr(BeginingPosition + Len(FirstLetter) + 1, Text, LastLetter)
    If BeginingPosition = 0 Or EndPosition = 0 Then Exit Function
    MidFromTargetCharToTargetChar = _
        IIf(ContainsFirstLetter, FirstLetter, vbNullString) & _
        Mid$(Text, BeginingPosition + Len(FirstLetter), EndPosition - BeginingPosition - Len(FirstLetter)) & _
        IIf(ContainsLastLetter, LastLetter, vbNullString)
End Function
Public Function NotIsArray(ByRef Target As Variant) As Boolean
    On Error Resume Next
    NotIsArray = Not (UBound(Target) > 0)
    NotIsArray = CBool(Err.Number <> 0)
    On Error GoTo 0
End Function
Public Sub OpenFolderExplorer(ByVal FolderPath As String)
    If Not FSO.FolderExists(FolderPath) Then Exit Sub
    Dim ExplorerPath(2) As String
    ExplorerPath(0) = "C:"
    ExplorerPath(1) = "Windows"
    ExplorerPath(2) = "explorer.exe"
    Dim Notes(1) As String
    Notes(0) = Join(ExplorerPath, AS_FOLDER)
    Notes(1) = FolderPath & AS_FOLDER
    Shell Join(Notes, HALF_SPACE), vbNormalFocus
End Sub
Public Sub QuickSortKeyAndItem(ByRef KeyArray() As Variant, ByRef ItemArray() As Variant, ByVal MinIndex As Long, ByVal MaxIndex As Long)
    Dim SmallIndex As Long, LargeIndex As Long, KeyMid As Variant
    SmallIndex = MinIndex
    LargeIndex = MaxIndex
    KeyMid = KeyArray(Int((SmallIndex + LargeIndex) / 2))
    Do
        Do While StrComp(KeyArray(SmallIndex), KeyMid) = -1
            SmallIndex = SmallIndex + 1
        Loop
        Do While StrComp(KeyArray(LargeIndex), KeyMid) = 1
            LargeIndex = LargeIndex - 1
        Loop
        If SmallIndex >= LargeIndex Then Exit Do
        Dim TemporaryKey As Variant
        TemporaryKey = KeyArray(SmallIndex)
        KeyArray(SmallIndex) = KeyArray(LargeIndex)
        KeyArray(LargeIndex) = TemporaryKey
        Dim TemporaryItem As Variant
        TemporaryItem = ItemArray(SmallIndex)
        ItemArray(SmallIndex) = ItemArray(LargeIndex)
        ItemArray(LargeIndex) = TemporaryItem
        SmallIndex = SmallIndex + 1
        LargeIndex = LargeIndex - 1
    Loop
    If (MinIndex < SmallIndex - 1) Then QuickSortKeyAndItem KeyArray, ItemArray, MinIndex, SmallIndex - 1
    If (MaxIndex > LargeIndex + 1) Then QuickSortKeyAndItem KeyArray, ItemArray, LargeIndex + 1, MaxIndex
End Sub
Public Function ReadAllLine(ByVal FilePath As String) As String
    On Error GoTo ThrowError
    With FSO.GetFile(FilePath).OpenAsTextStream
        ReadAllLine = .ReadAll
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Function
Public Function ReadOneLineAsArray1D(ByVal FilePath As String) As String()
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForReading)
        Dim Text As String
        Text = .ReadLine
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
    ReadOneLineAsArray1D = Split(Text, CANNMA)
End Function
Public Function ReadAllLineAsArray2D(ByVal FilePath As String) As Variant()
    Dim Headers() As String
    Headers = LIB.ReadOneLineAsArray1D(FilePath)
    If LIB.NotIsArray(Headers) Then Exit Function
    On Error GoTo ThrowError
    Dim MaxR As Long, MaxC As Long
    MaxR = LIB.ReadLineCount(FilePath)
    MaxC = UBound(Headers) - LBound(Headers) + 1
    Dim RET() As Variant
    ReDim RET(1 To MaxR, 1 To MaxC)
    With FSO.OpenTextFile(FilePath, ForReading)
        Do Until .AtEndOfLine
            Dim Texts() As String
            Texts = Split(.ReadLine, CANNMA)
            Dim R As Long, C As Long
            R = R + 1
            For C = 1 To MaxC
                RET(R, C) = Texts(C - 1)
            Next
        Loop
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
    ReadAllLineAsArray2D = RET
End Function
Public Function ReadLineCount(ByVal FilePath As String) As Long
    On Error GoTo ThrowError
    Dim DataBody() As Byte
    Open FilePath For Binary As #1
        ReDim DataBody(1 To LOF(1))
        Get #1, , DataBody
        Dim Text As String
        Text = StrConv(DataBody, vbUnicode)
ThrowError:
    Close #1
    DoEvents
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
    ReadLineCount = UBound(Split(Text, vbLf))
End Function
Public Sub RemoveSelectedItemOfListBox(ByVal ListBoxObject As MSForms.ListBox)
    With ListBoxObject
        Dim i As Long
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then .RemoveItem i
        Next
    End With
End Sub
Public Function ReplaceIllegalCharacter(ByVal Text As String, ParamArray Letters()) As String
    If Text Like vbNullString Then Exit Function
    If UBound(Letters) = -1 Then Exit Function
    Dim RET As String
    RET = Text
    Dim i As Long
    For i = LBound(Letters) To UBound(Letters)
        If InStr(RET, Letters(i)) Then
            RET = Replace(RET, Letters(i), vbNullString)
        End If
    Next
    ReplaceIllegalCharacter = RET
End Function
Public Function ReplaceLineBreakChar→VbLf(ByVal Text As Variant, ByVal LineBreakChar As String) As String
    ReplaceLineBreakChar→VbLf = Replace(Text, LineBreakChar, vbLf)
End Function
Public Function ReplaceVbLf→LineBreakChar(ByVal Text As Variant, ByVal LineBreakChar As String) As String
    ReplaceVbLf→LineBreakChar = Replace(Text, vbLf, LineBreakChar)
End Function
Public Function Replace○曜日→WeekDay(ByVal ○曜日 As String) As Long
    Dim RET As Long
    Select Case ○曜日
        Case "月曜日"
            RET = vbMonday
        Case "火曜日"
            RET = vbTuesday
        Case "水曜日"
            RET = vbWednesday
        Case "木曜日"
            RET = vbThursday
        Case "金曜日"
            RET = vbFriday
        Case "土曜日"
            RET = vbSaturday
        Case "日曜日"
            RET = vbSunday
    End Select
    Replace○曜日→WeekDay = RET
End Function
Public Function ReturnArrayDimension(ByVal Target As Variant) As Long
    On Error GoTo ExitLoop
    Do
        Dim Dimensions As Long, ArrayElements As Long
        Dimensions = Dimensions + 1
        ArrayElements = UBound(Target, Dimensions)
    Loop
ExitLoop:
    On Error GoTo 0
    ReturnArrayDimension = Dimensions - 1 + ArrayElements - ArrayElements
End Function
Public Function ReturnDateByInputBox(ByVal Promput As String, ByVal Title As String, ByVal Default As Date) As String
    Dim RET As String
    Do
        RET = InputBox(Promput, Title, Default)
        If RET Like vbNullString Then Exit Function
        If IsDate(RET) And CDate(RET) <> 0 Then Exit Do
        MsgBox "日付の入力形式が正しくありません。", vbExclamation, Title
    Loop
    ReturnDateByInputBox = RET
End Function
Public Function ReturnVbColor(ByVal TargetDate As Date, ByVal NomalColor As Long, ByVal SatColor As Long, ByVal SunColor As Long) As Long
    Dim RET As Long
    Select Case Weekday(TargetDate)
        Case vbMonday To vbFriday
            RET = NomalColor
        Case vbSaturday
            RET = SatColor
        Case vbSunday
            RET = SunColor
    End Select
    ReturnVbColor = RET
End Function
Public Function SelectedItemJustOne(ByVal Target As MSForms.ListBox) As String
    With Target
        Dim i As Long
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Dim RET As String
                RET = .List(i)
                Dim SelectedLists As Long
                SelectedLists = SelectedLists + 1
            End If
        Next
    End With
    If SelectedLists = 1 Then SelectedItemJustOne = RET
End Function
Public Function SortDictionary(ByVal Target As Dictionary, Optional ByVal Order As SORT_ORDER = xlAscending) As Dictionary
    If Target Is Nothing Then Exit Function
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim TargetKeys() As Variant
    TargetKeys = Target.Keys
    TargetKeys = MergeSort(TargetKeys, Order)
    Dim TargetKey As Variant
    For Each TargetKey In TargetKeys
        RET.Add TargetKey, Target.Item(TargetKey)
    Next
    Set SortDictionary = RET
End Function
Public Function ToAbsolute(ByVal Relative As String) As String
    If Relative Like vbNullString Then Exit Function
    Dim RET As String
    RET = Application.ConvertFormula("=" & Relative, xlA1, xlA1, ToAbsolute:=xlAbsolute)
    RET = Replace(RET, "=", vbNullString)
    ToAbsolute = RET
End Function
Public Function ToAlphabetNotation(ByVal WesternYear As Long) As String
    ToAlphabetNotation = Format$(DateSerial(WesternYear, 1, 1), "ge")
End Function
Public Function ToColumnIndex(ByVal CellAddress As String) As Long
    Dim RET As Long
    With New RegExp
        .Pattern = "[0-9,$]*"
        .Global = True
        Dim ColumnLetter As String
        ColumnLetter = .Replace(CellAddress, vbNullString)
    End With
    Dim Place As Long
    For Place = Len(ColumnLetter) To 1 Step -1
        Dim AlphabetCode As Long
        AlphabetCode = asc(Mid$(ColumnLetter, Place, 1)) - 64
        If Len(ColumnLetter) - Place <> 0 Then
            RET = RET + AlphabetCode * 26 ^ (Len(ColumnLetter) - Place)
        Else
            RET = AlphabetCode
        End If
    Next
    If RET > 16384 Then MsgBox "列番号の最大値を超えています。", vbExclamation, "オーバーフロー": Exit Function
    ToColumnIndex = RET
End Function
Public Function ToRowIndex(ByVal CellAddress As String) As Long
    With New RegExp
        .Pattern = "[A-Z,$]*"
        .Global = True
        ToRowIndex = .Replace(CellAddress, vbNullString)
    End With
End Function
Public Function TransposeArray1D→2D(ByRef Array1D As Variant) As Variant()
    If NotIsArray(Array1D) Then Exit Function
    Dim Array2D() As Variant
    ReDim Array2D(1 To 1, 1 To UBound(Array1D) + 1)
    Dim i As Long
    For i = LBound(Array1D) To UBound(Array1D)
        Array2D(1, i + 1) = Array1D(i)
    Next
    TransposeArray1D→2D = Array2D
End Function
Public Function TransposeArray2D→1D(ByRef Array2D() As Variant, Optional ByVal BaseOne As Boolean) As Variant()
    If NotIsArray(Array2D) Then Exit Function
    If LIB.ReturnArrayDimension(Array2D) <> 2 Then Exit Function
    If UBound(Array2D, 1) > 1 And UBound(Array2D, 2) > 1 Then Exit Function
    Dim UpperIndex As Long
    If UBound(Array2D, 1) > 1 Then UpperIndex = UBound(Array2D, 1)
    If UBound(Array2D, 2) > 1 Then UpperIndex = UBound(Array2D, 2)
    Dim AddIndex As Long
    If BaseOne Then AddIndex = 1
    Dim Array1D() As Variant
    ReDim Array1D(0 + AddIndex To UpperIndex - 1 + AddIndex)
    Dim i As Long
    For i = 1 To UpperIndex
        If UBound(Array2D, 1) > 1 Then Array1D(i - 1 + AddIndex) = Array2D(i, 1)
        If UBound(Array2D, 2) > 1 Then Array1D(i - 1 + AddIndex) = Array2D(1, i)
    Next
    TransposeArray2D→1D = Array1D
End Function
Public Function TransposeCommonItem→Key(ByVal TargetDictionary As Dictionary) As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim TargetKey As Variant
    For Each TargetKey In TargetDictionary
        Dim TargetItem As Variant
        TargetItem = TargetDictionary.Item(TargetKey)
        If RET.Exists(TargetItem) Then
            Dim CollectionKey As Collection
            Set CollectionKey = RET.Item(TargetItem)
            CollectionKey.Add TargetKey
            Set RET.Item(TargetItem) = CollectionKey
        Else
            Set CollectionKey = New Collection
            CollectionKey.Add TargetKey
            RET.Add TargetItem, CollectionKey
        End If
    Next
    Set TransposeCommonItem→Key = RET
End Function
Public Function TrimVbLf(ByVal Text As String) As String
    With New RegExp
        .Pattern = "^\n+|\n+$|\n+(?=\n)"
        .Global = True
        TrimVbLf = .Replace(Text, vbNullString)
    End With
End Function
Public Function VbLfToLineBreakChar(ByVal Text As String) As String
    Dim RET As String
    If InStr(Text, vbLf) > 0 Then
        RET = Replace(Text, vbLf, LINE_BREAK_CHAR)
    Else
        RET = Text
    End If
    VbLfToLineBreakChar = RET
End Function
Public Sub WriteOneLineForAppending(ByVal FilePath As String, ByVal Text As String)
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForAppending, True)
        .WriteLine Text
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Sub
Public Sub WriteOneLine(ByVal FilePath As String, ByVal Text As String)
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        .WriteLine Text
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Sub
Public Sub WriteSomeLines(ByVal FilePath As String, ByRef DataBody2D() As Variant)
    Dim Texts() As Variant
    ReDim Texts(LBound(DataBody2D, 2) To UBound(DataBody2D, 2))
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        Dim R As Long, C As Long
        For R = LBound(DataBody2D, 1) To UBound(DataBody2D, 1)
            For C = LBound(DataBody2D, 2) To UBound(DataBody2D, 2)
                Texts(C) = DataBody2D(R, C)
            Next
            Dim Text As String
            Text = Join(Texts, CANNMA)
            .WriteLine Text
        Next
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Sub
Public Function WesternYear(ByVal TargetDate As Date) As Long
    If TargetDate = 0 Then Exit Function
    WesternYear = IIf(Month(TargetDate) < 4, Year(TargetDate) - 1, Year(TargetDate))
End Function


