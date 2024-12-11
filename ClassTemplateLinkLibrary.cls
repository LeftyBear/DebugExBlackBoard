VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTemplateLinkLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Template")
Option Explicit
Private Sheet_                  As Worksheet
Private UsedPrintingRange_      As Range
Private RecordNumbers_          As Dictionary
Private PrintingItemIndexes_    As Dictionary
Private ArrayPositions_         As Dictionary
Private ScheduleValues_         As Dictionary
Private PrintingItem_           As ClassPrintingItem
Private DailySchedule_          As ClassDailySchedule
Private WeeklySchedule_         As ClassWeeklySchedule
Private MonthlySchedule_        As ClassMonthlySchedule
Private YearlySchedule_         As ClassYearlySchedule
Private MaxFontSize_            As Long
Private MinFontSize_            As Long
Private ReplacedVbLf_           As String
Private Const DIVISOR_ROW_HEIGHT    As Long = 2
Private Const HATCHING              As Long = xlPatternGray25
Private Type PrintingItemHeaders
    '@Ignore UseMeaningfulName
    €–Ú–¼      As Long
    '@Ignore UseMeaningfulName
    €–Ú–¼’Ç‰Á  As Long
    '@Ignore UseMeaningfulName
    ˆóü”ÍˆÍ    As Long
    '@Ignore UseMeaningfulName
    ˆóüT      As Long
    '@Ignore UseMeaningfulName
    WŒv•û–@    As Long
    '@Ignore UseMeaningfulName
    ˆóüŒŽ      As Long
End Type
Private Type PrintingItemVariable
    ItemName            As String
    VisibledItemName    As Boolean
    PrintingAddress     As String
    WeekName            As String
    TotalType           As String
    MonthName           As String
End Type
Public Enum SCHEDULE_TYPE
    Daily = 1
    Weekly
    Monthly
    Yearly
End Enum
Private Const MODULE_NAME As String = "ClassTemplateLinkLibrary"
Public Property Get Sheet() As Worksheet
    Set Sheet = Sheet_
End Property
Public Property Set Sheet(ByVal Target As Worksheet)
    Set Sheet_ = Target
End Property
Public Property Get UsedPrintingRange() As Range
    Set UsedPrintingRange = UsedPrintingRange_
End Property
Public Property Set UsedPrintingRange(ByVal Target As Range)
    Set UsedPrintingRange_ = Target
End Property
Public Property Get RecordNumbers() As Dictionary
    Set RecordNumbers = RecordNumbers_
End Property
Public Property Set RecordNumbers(ByVal Target As Dictionary)
    Set RecordNumbers_ = Target
End Property
Public Property Get PrintingItemIndexes() As Dictionary
    Set PrintingItemIndexes = PrintingItemIndexes_
End Property
Public Property Set PrintingItemIndexes(ByVal Target As Dictionary)
    Set PrintingItemIndexes_ = Target
End Property
Public Property Get ArrayPositions() As Dictionary
    Set ArrayPositions = ArrayPositions_
End Property
Public Property Set ArrayPositions(ByVal Target As Dictionary)
    Set ArrayPositions_ = Target
End Property
Public Property Get ScheduleValues() As Dictionary
    Set ScheduleValues = ScheduleValues_
End Property
Public Property Set ScheduleValues(ByVal Target As Dictionary)
    Set ScheduleValues_ = Target
End Property
Public Property Get PrintingItem() As ClassPrintingItem
    Set PrintingItem = PrintingItem_
End Property
Public Property Set PrintingItem(ByVal Target As ClassPrintingItem)
    Set PrintingItem_ = Target
End Property
Public Property Get DailySchedule() As ClassDailySchedule
    Set DailySchedule = DailySchedule_
End Property
Public Property Set DailySchedule(ByVal Target As ClassDailySchedule)
    Set DailySchedule_ = Target
End Property
Public Property Get WeeklySchedule() As ClassWeeklySchedule
    Set WeeklySchedule = WeeklySchedule_
End Property
Public Property Set WeeklySchedule(ByVal Target As ClassWeeklySchedule)
    Set WeeklySchedule_ = Target
End Property
Public Property Get MonthlySchedule() As ClassMonthlySchedule
    Set MonthlySchedule = MonthlySchedule_
End Property
Public Property Set MonthlySchedule(ByVal Target As ClassMonthlySchedule)
    Set MonthlySchedule_ = Target
End Property
Public Property Get YearlySchedule() As ClassYearlySchedule
    Set YearlySchedule = YearlySchedule_
End Property
Public Property Set YearlySchedule(ByVal Target As ClassYearlySchedule)
    Set YearlySchedule_ = Target
End Property
Public Property Let MaxFontSize(ByVal Target As Long)
    MaxFontSize_ = Target
End Property
Public Property Get MaxFontSize() As Long
    MaxFontSize = MaxFontSize_
End Property
Public Property Let MinFontSize(ByVal Target As Long)
    MinFontSize_ = Target
End Property
Public Property Get MinFontSize() As Long
    MinFontSize = MinFontSize_
End Property
Public Property Let ReplacedVbLf(ByVal Target As String)
    ReplacedVbLf_ = Target
End Property
Public Property Get ReplacedVbLf() As String
    ReplacedVbLf = ReplacedVbLf_
End Property
Public Function AddUsedPrintingRange() As Range
    Dim RET As Range
    Dim BufSheet As Worksheet
    Set BufSheet = Me.Sheet
    Dim BufDataBody() As Variant
    '@Ignore UseMeaningfulName
    Dim Cˆóü”ÍˆÍ As Long
    With Me.PrintingItem
        BufDataBody = .DataBody
        Cˆóü”ÍˆÍ = .Cˆóü”ÍˆÍ
    End With
    Dim R As Long
    For R = 1 To UBound(BufDataBody)
        Dim PrintingRange As Range
        Set PrintingRange = BufSheet.Range(CStr(BufDataBody(R, Cˆóü”ÍˆÍ)))
        If Not RET Is Nothing Then
            Set RET = BufSheet.Range(RET, PrintingRange)
        Else
            Set RET = PrintingRange
        End If
    Next
    Set AddUsedPrintingRange = RET
End Function
Public Function AddRecordNumbers(ByVal RecordNumbers As Collection) As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim BufSheet As Worksheet
    Set BufSheet = Me.Sheet
    Dim BufDataBody() As Variant
    '@Ignore UseMeaningfulName
    Dim Cˆóü”ÍˆÍ As Long
    '@Ignore UseMeaningfulName
    Dim CˆóüT As Long
    '@Ignore UseMeaningfulName
    Dim CˆóüŒŽ As Long
    With Me.PrintingItem
        BufDataBody = .DataBody
        Cˆóü”ÍˆÍ = .Cˆóü”ÍˆÍ
        CˆóüT = .CˆóüT
        CˆóüŒŽ = .CˆóüŒŽ
    End With
    Dim R As Long
    For R = 1 To UBound(BufDataBody)
        Dim PrintingRange As Range
        Dim TargetRange As Range
        Dim TargetAddress As String
        Set PrintingRange = BufSheet.Range(BufDataBody(R, Cˆóü”ÍˆÍ))
        Dim i As Long
        Dim RecordNumber As Long
        i = ReturnRecordNumbersIndex(BufDataBody(R, CˆóüT), BufDataBody(R, CˆóüŒŽ))
        If IsSingleCell(PrintingRange) Then
            Set TargetRange = PrintingRange
            TargetAddress = TargetRange.Address(False, False)
            RecordNumber = RecordNumbers.Item(i)
            If Not RET.Exists(TargetAddress) Then
                RET.Add TargetAddress, RecordNumber
            End If
        ElseIf IsContinuousSingleCell(PrintingRange) Then
            For Each TargetRange In PrintingRange
                TargetAddress = TargetRange.Address(False, False)
                RecordNumber = RecordNumbers.Item(i)
                If Not RET.Exists(TargetAddress) Then
                    RET.Add TargetAddress, RecordNumber
                End If
                i = i + 1
            Next
        ElseIf HasMergeCell(PrintingRange) Then
            For Each TargetRange In PrintingRange
                Dim UpperLeftCell As Range
                Set UpperLeftCell = TargetRange.MergeArea.Item(1)
                If TargetRange.Address Like UpperLeftCell.Address Then
                    TargetAddress = TargetRange.Address(False, False)
                    RecordNumber = RecordNumbers.Item(i)
                    If Not RET.Exists(TargetAddress) Then
                        RET.Add TargetAddress, RecordNumber
                    End If
                    i = i + 1
                End If
            Next
        End If
    Next
    Set AddRecordNumbers = RET
End Function
Public Function AddPrintingItemIndexes() As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim BufSheet As Worksheet
    Set BufSheet = Me.Sheet
    Dim BufDataBody() As Variant
    '@Ignore UseMeaningfulName
    Dim Cˆóü”ÍˆÍ As Long
    With Me.PrintingItem
        BufDataBody = .DataBody
        Cˆóü”ÍˆÍ = .Cˆóü”ÍˆÍ
    End With
    Dim R As Long
    For R = 1 To UBound(BufDataBody)
        Dim PrintingRange As Range
        Dim TargetRange As Range
        Dim TargetAddress As String
        Set PrintingRange = BufSheet.Range(BufDataBody(R, Cˆóü”ÍˆÍ))
        If IsSingleCell(PrintingRange) Then
            Set TargetRange = PrintingRange
            TargetAddress = TargetRange.Address(False, False)
            If RET.Exists(TargetAddress) Then
                RET.Item(TargetAddress) = RET.Item(TargetAddress) & CANNMA & R
            Else
                RET.Add TargetAddress, R
            End If
        ElseIf IsContinuousSingleCell(PrintingRange) Then
            For Each TargetRange In PrintingRange
                TargetAddress = TargetRange.Address(False, False)
                If RET.Exists(TargetAddress) Then
                    RET.Item(TargetAddress) = RET.Item(TargetAddress) & CANNMA & R
                Else
                    RET.Add TargetAddress, R
                End If
            Next
        ElseIf HasMergeCell(PrintingRange) Then
            For Each TargetRange In PrintingRange
                Dim UpperLeftCell As Range
                Set UpperLeftCell = TargetRange.MergeArea.Item(1)
                If TargetRange.Address Like UpperLeftCell.Address Then
                    TargetAddress = TargetRange.Address(False, False)
                    If RET.Exists(TargetAddress) Then
                        RET.Item(TargetAddress) = RET.Item(TargetAddress) & CANNMA & R
                    Else
                        RET.Add TargetAddress, R
                    End If
                End If
            Next
        End If
    Next
    Set AddPrintingItemIndexes = RET
End Function
Public Function AddArrayPositions() As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim BufSheet As Worksheet
    Dim BufUsedPrintingRange As Range
    Dim BufPrintingItemIndexes As Dictionary
    Set BufSheet = Me.Sheet
    Set BufUsedPrintingRange = Me.UsedPrintingRange
    Set BufPrintingItemIndexes = Me.PrintingItemIndexes
    Dim PrintingAddress As Variant
    For Each PrintingAddress In BufPrintingItemIndexes
        Dim PrintingRange As Range
        Set PrintingRange = BufSheet.Range(PrintingAddress)
        If Not RET.Exists(PrintingAddress) Then
            Dim CellDiffAddress As String
            CellDiffAddress = LIB.CellDiff(BufSheet, BufUsedPrintingRange, PrintingRange)
            RET.Add PrintingAddress, CellDiffAddress
        End If
    Next
    Set AddArrayPositions = RET
End Function
Public Function AddScheduleValues(ByVal ScheduleType As SCHEDULE_TYPE) As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim BufPrintingItemIndexes As Dictionary
    Dim BufRecordNumbers As Dictionary
    Set BufPrintingItemIndexes = Me.PrintingItemIndexes
    Set BufRecordNumbers = Me.RecordNumbers
    Dim BufDataBody() As Variant
    Dim C As PrintingItemHeaders
    BufDataBody = Me.PrintingItem.DataBody
    With C
        .€–Ú–¼ = CLng(Me.PrintingItem.C€–Ú–¼)
        .€–Ú–¼’Ç‰Á = CLng(Me.PrintingItem.C€–Ú–¼’Ç‰Á)
        .ˆóü”ÍˆÍ = CLng(Me.PrintingItem.Cˆóü”ÍˆÍ)
        .ˆóüT = CLng(Me.PrintingItem.CˆóüT)
        .WŒv•û–@ = CLng(Me.PrintingItem.CWŒv•û–@)
        .ˆóüŒŽ = CLng(Me.PrintingItem.CˆóüŒŽ)
    End With
    Dim PrintingAddress As Variant
    For Each PrintingAddress In BufPrintingItemIndexes
        Dim RecordNumber As Long
        Dim Indexes() As String
        RecordNumber = BufRecordNumbers.Item(PrintingAddress)
        Indexes = Split(PrintingItemIndexes.Item(PrintingAddress), CANNMA)
        Dim i As Long
        For i = LBound(Indexes) To UBound(Indexes)
            Dim R As Long
            R = Indexes(i)
            Dim PrintingItems As PrintingItemVariable
            With PrintingItems
                .ItemName = BufDataBody(R, C.€–Ú–¼)
                .VisibledItemName = IIf(BufDataBody(R, C.€–Ú–¼’Ç‰Á) Like "‚·‚é", True, False)
                .PrintingAddress = BufDataBody(R, C.ˆóü”ÍˆÍ)
                .WeekName = BufDataBody(R, C.ˆóüT)
                .TotalType = BufDataBody(R, C.WŒv•û–@)
                .MonthName = BufDataBody(R, C.ˆóüŒŽ)
            End With
            Dim ScheduleValue As Variant
            ScheduleValue = ReturnScheduleValue(ScheduleType, RecordNumber, PrintingItems)
            If Not IsEmpty(ScheduleValue) Then
                If Not IsEmpty(RET.Item(PrintingAddress)) Then
                    RET.Item(PrintingAddress) = RET.Item(PrintingAddress) & vbLf & ScheduleValue
                Else
                    RET.Item(PrintingAddress) = ScheduleValue
                End If
            Else
                If Not RET.Exists(PrintingAddress) Then
                    RET.Add PrintingAddress, Empty
                End If
            End If
        Next
    Next
    Set AddScheduleValues = RET
End Function
Public Function AddDataBody() As Variant()
    Dim RET() As Variant
    RET = AddDataBodyWithFormula
    Dim BufSheet As Worksheet
    Dim BufArrayPositions As Dictionary
    Dim BufScheduleValues As Dictionary
    Set BufSheet = Me.Sheet
    Set BufArrayPositions = Me.ArrayPositions
    Set BufScheduleValues = Me.ScheduleValues
    Dim PrintingAddress As Variant
    For Each PrintingAddress In BufArrayPositions
        If Not IsEmpty(BufScheduleValues.Item(PrintingAddress)) Then
            Dim R As Long
            Dim C As Long
            R = BufSheet.Range(BufArrayPositions.Item(PrintingAddress)).Row
            C = BufSheet.Range(BufArrayPositions.Item(PrintingAddress)).Column
            RET(R, C) = BufScheduleValues.Item(PrintingAddress)
        End If
    Next
    AddDataBody = RET
End Function
Public Sub Size•¶ŽšƒTƒCƒYk¬()
    On Error GoTo ThrowError
    Const ProcedureName As String = "Size•¶ŽšƒTƒCƒYk¬"
    Dim Addresses As Dictionary
    Set Addresses = AddFontSizeDownAddresses
    If EnabledReplaceVbLf Then ReplaceVbLf Addresses
    Dim DownSizes As Dictionary
    Set DownSizes = LIB.TransposeCommonItem¨Key(Addresses)
    Dim BufSheet As Worksheet
    Set BufSheet = Me.Sheet
    Dim DownSize As Variant
    For Each DownSize In DownSizes
        Dim TargetAddress As Variant
        For Each TargetAddress In DownSizes.Item(DownSize)
            Dim UnionTarget As Range
            If Not UnionTarget Is Nothing Then
                Set UnionTarget = Union(UnionTarget, BufSheet.Range(TargetAddress))
            Else
                Set UnionTarget = BufSheet.Range(TargetAddress)
            End If
        Next
        UnionTarget.Font.Size = DownSize
        Set UnionTarget = Nothing
    Next
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub Toc‘‚«‚Ì“ú•t()
    On Error GoTo ThrowError
    Const ProcedureName As String = "Toc‘‚«‚Ì“ú•t"
    With Me.PrintingItem
        Dim i As Long
        For i = 1 To UBound(.DataBody)
            Dim PrintingRange As Range
            Set PrintingRange = Me.Sheet.Range(.Itemˆóü”ÍˆÍ(i))
            If IsDateContinuousCell(PrintingRange, .Item€–Ú–¼(i)) Then
                Dim Target As Range
                Dim UpperLeft As Range
                For Each Target In PrintingRange
                    If Target.MergeArea.Height > Target.MergeArea.Width Then
                        Set UpperLeft = Target.MergeArea.Item(1)
                        If Target.Address Like UpperLeft.Address Then
                            If IsValidDate(Target.Value) Then
                                Target.Value = ToVerticalWritingDate(Target.Value)
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End With
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub Color“ú•t’PˆêƒZƒ‹by“ú’¼“ú()
    On Error GoTo ThrowError
    Const ProcedureName As String = "Color“ú•t’PˆêƒZƒ‹by“ú’¼“ú"
    Dim DayDutyDays As Dictionary
    Set DayDutyDays = Schedule.TableColumnItemList("“ú’¼“ú")
    With Me.PrintingItem
        Dim i As Long
        For i = 1 To UBound(.DataBody)
            If .Item€–Ú–¼(i) Like "“ú•t" Then
                Dim Target As Range
                Set Target = Me.Sheet.Range(.Itemˆóü”ÍˆÍ(i))
                If IsValidDate(Target.Value) Then
                    If CBool(DayDutyDays.Item(CDate(Target.Value))) Then
                        Target.Font.Color = RGB_COLOR.Black
                    Else
                        Target.Font.Color = RGB_COLOR.Red
                    End If
                End If
            End If
        Next
    End With
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub Pattern–ÔŠ|‚¯by“ú’¼“ú(ByVal UsedPrintingRange As Range)
    On Error GoTo ThrowError
    Const ProcedureName As String = "Pattern–ÔŠ|‚¯by“ú’¼“ú"
    Dim C As Long
    C = UsedPrintingRange.Columns.Count
    Dim BufSheet As Worksheet
    Set BufSheet = Me.Sheet
    Dim DayDutyDays As Dictionary
    Set DayDutyDays = Schedule.TableColumnItemList("“ú’¼“ú")
    With Me.PrintingItem
        Dim i As Long
        For i = 1 To UBound(.DataBody)
            Dim PrintingRange As Range
            Set PrintingRange = BufSheet.Range(.Itemˆóü”ÍˆÍ(i))
            If IsDateContinuousCell(PrintingRange, .Item€–Ú–¼(i)) Then
                Dim Target As Range
                For Each Target In PrintingRange
                    If IsValidDate(Target.Value) Then
                        If Not CBool(DayDutyDays.Item(CDate(Target.Value))) Then
                            Dim R As Long
                            Dim UnionTarget As Range
                            R = Target.Row
                            If Not UnionTarget Is Nothing Then
                                Set UnionTarget = Union(UnionTarget, BufSheet.Range(Target, BufSheet.Cells.Item(R, C)))
                            Else
                                Set UnionTarget = BufSheet.Range(Target, BufSheet.Cells.Item(R, C))
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End With
    If Not UnionTarget Is Nothing Then UnionTarget.Interior.Pattern = HATCHING
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub Pattern–ÔŠ|‚¯byŽö‹Æ“ú()
    On Error GoTo ThrowError
    Const ProcedureName As String = "Pattern–ÔŠ|‚¯byŽö‹Æ“ú"
    Dim BufSheet As Worksheet
    Set BufSheet = Me.Sheet
    Dim ClassDays As Dictionary
    Set ClassDays = Schedule.TableColumnItemList("Žö‹Æ“ú")
    If ClassDays Is Nothing Then Exit Sub
    With Me.PrintingItem
        Dim i As Long
        For i = 1 To UBound(.DataBody)
            Dim PrintingRange As Range
            Dim ItemName As String
            Set PrintingRange = BufSheet.Range(.Itemˆóü”ÍˆÍ(i))
            ItemName = .Item€–Ú–¼(i)
            If IsDateContinuousCell(PrintingRange, ItemName) Then
                Dim Target As Range
                Dim UnionTarget As Range
                For Each Target In PrintingRange
                    If IsValidDate(Target.Value) Then
                        If Not CBool(ClassDays.Item(CDate(Target.Value))) Then
                            If Not UnionTarget Is Nothing Then
                                Set UnionTarget = Union(UnionTarget, Target)
                            Else
                                Set UnionTarget = Target
                            End If
                        End If
                    End If
                Next
            ElseIf ItemName Like "‘åsŽ–" Then
                For Each Target In PrintingRange
                    Dim OffsetTarget As Range
                    Set OffsetTarget = Target.Offset(0, -1)
                    If WorksheetFunction.IsNumber(OffsetTarget.Value) Then
                        If IsDate(OffsetTarget.Value) Or IsNumeric(OffsetTarget.Value) Then
                            If Not ClassDays.Item(CDate(OffsetTarget.Value)) Then
                                If Not UnionTarget Is Nothing Then
                                    Set UnionTarget = Union(UnionTarget, Target)
                                Else
                                    Set UnionTarget = Target
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End With
    If Not UnionTarget Is Nothing Then UnionTarget.Interior.Pattern = HATCHING
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub Fixs‚by“ú’¼“ú()
    On Error GoTo ThrowError
    Const ProcedureName As String = "Fixs‚by“ú’¼“ú"
    Const BasisRowHeightAddress As String = "A5"
    Dim BufSheet As Worksheet
    Set BufSheet = Me.Sheet
    Dim BasisRowHeight As Double
    BasisRowHeight = BufSheet.Range(BasisRowHeightAddress).RowHeight
    Dim DayDutyDays As Dictionary
    Set DayDutyDays = Schedule.TableColumnItemList("“ú’¼“ú")
    If DayDutyDays Is Nothing Then Exit Sub
    With Me.PrintingItem
        Dim i As Long
        For i = 1 To UBound(.DataBody)
            Dim PrintingRange As Range
            Set PrintingRange = BufSheet.Range(.Itemˆóü”ÍˆÍ(i))
            If IsDateContinuousCell(PrintingRange, .Item€–Ú–¼(i)) Then
                Dim Target As Range
                Dim UnionTarget As Range
                For Each Target In PrintingRange
                    If Not IsEmpty(Target.Value) Then
                        If IsValidDate(Target.Value) Then
                            If Not CBool(DayDutyDays.Item(CDate(Target.Value))) Then
                                If Not UnionTarget Is Nothing Then
                                    Set UnionTarget = Union(UnionTarget, Target)
                                Else
                                    Set UnionTarget = Target
                                End If
                            End If
                        End If
                    Else
                        Target.EntireRow.Hidden = True
                    End If
                Next
                Exit For
            End If
        Next
    End With
    If Not UnionTarget Is Nothing Then UnionTarget.RowHeight = BasisRowHeight / DIVISOR_ROW_HEIGHT
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Function HasSheet(ByVal Sheet As Worksheet, ByVal MsgTitle As String) As Boolean
    If Sheet Is Nothing Then
        MsgBox "ˆóü—lŽ®ƒtƒ@ƒCƒ‹‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñ‚Å‚µ‚½B", vbExclamation, MsgTitle: Exit Function
    End If
    HasSheet = True
End Function
Public Function HasPrintingItem(ByVal PrintingItem As ClassPrintingItem, ByVal MsgTitle As String) As Boolean
    If LIB.NotIsArray(PrintingItem.DataBody) Then
        MsgBox "ˆóü—lŽ®Ý’èƒtƒ@ƒCƒ‹‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñ‚Å‚µ‚½B", vbExclamation, MsgTitle: Exit Function
    End If
    HasPrintingItem = True
End Function
Private Function AddDataBodyWithFormula() As Variant()
    Dim RET() As Variant
    RET = Me.UsedPrintingRange.Value
    Dim BufUsedPrintingRange As Range
    Set BufUsedPrintingRange = Me.UsedPrintingRange
    Dim R As Long
    Dim C As Long
    For R = 1 To UBound(RET, 1)
        For C = 1 To UBound(RET, 2)
            Dim Target As Range
            Set Target = BufUsedPrintingRange.Cells.Item(R, C)
            If Target.HasFormula Then
                RET(R, C) = Target.Formula
            End If
        Next
    Next
    AddDataBodyWithFormula = RET
End Function
Private Function ReturnRecordNumbersIndex(ByVal WeekName As String, ByVal MonthName As String) As Long
    Dim RET As Long
    If Not WeekName Like vbNullString Then
        Select Case True
            Case WeekName Like "¡T"
                RET = 1
            Case WeekName Like "—ˆT"
                RET = 6
        End Select
    ElseIf Not MonthName Like vbNullString Then
        Select Case CLng(MonthName)
            Case 1 To 3
                RET = (CLng(MonthName) + 8) * 31 + 1
            Case 4 To 12
                RET = (CLng(MonthName) - 4) * 31 + 1
        End Select
    Else
        RET = 1
    End If
    ReturnRecordNumbersIndex = RET
End Function
Private Function ReturnScheduleValue(ByVal ScheduleType As SCHEDULE_TYPE, ByVal RecordNumber As Long, ByRef PrintingItems As PrintingItemVariable) As Variant
    Dim RET As Variant
    With PrintingItems
        Select Case ScheduleType
            Case SCHEDULE_TYPE.Daily
                RET = ReturnDailyScheduleValue(RecordNumber, .ItemName, .VisibledItemName)
            Case SCHEDULE_TYPE.Weekly
                RET = ReturnWeeklyScheduleValue(RecordNumber, .ItemName, .VisibledItemName, .PrintingAddress, .WeekName, .TotalType)
            Case SCHEDULE_TYPE.Monthly
                RET = ReturnMonthlyScheduleValue(RecordNumber, .ItemName, .VisibledItemName)
            Case SCHEDULE_TYPE.Yearly
                RET = ReturnYearlyScheduleValue(RecordNumber, .ItemName, .VisibledItemName, .MonthName)
        End Select
    End With
    ReturnScheduleValue = RET
End Function
Private Function ReturnDailyScheduleValue(ByVal RecordNumber As Long, ByVal ItemName As String, ByVal VisibledItemName As Boolean) As Variant
    If RecordNumber = 0 Then Exit Function
    Dim RET As Variant
    With Me.DailySchedule
        RET = .DataBody(RecordNumber, .ItemNames.Item(ItemName))
    End With
    If IsEmpty(RET) Then Exit Function
    If VisibledItemName Then
        If InStr(RET, vbLf) Then
            Dim Texts() As String
            Texts = Split(RET, vbLf)
            Dim i As Long
            For i = LBound(Texts) To UBound(Texts)
                Texts(i) = "y" & ItemName & "z" & Texts(i)
            Next
            RET = Join(Texts, vbLf)
        Else
            RET = "y" & ItemName & "z" & RET
        End If
    End If
    ReturnDailyScheduleValue = RET
End Function
Private Function ReturnWeeklyScheduleValue(ByVal RecordNumber As Long, ByVal ItemName As String, ByVal VisibledItemName As Boolean, _
ByVal PrintingAddress As String, ByVal WeekName As String, ByVal TotalType As String) As Variant
    Dim RET As Variant
    Select Case True
        Case WeekName Like "¡T"
            Select Case True
                Case IsDateSingleCell(ItemName, PrintingAddress)
                    RET = Me.WeeklySchedule.ThisWeekFirstDate
                Case ItemName Like "T”"
                    RET = Me.WeeklySchedule.ThisWeekClassWeeks
                Case TotalType Like "TŒv"
                    RET = Me.WeeklySchedule.ThisWeekTotalClassHourOfWeek(ItemName)
                Case TotalType Like "—ÝŒv"
                    RET = Me.WeeklySchedule.ThisWeekTotalClassHourOfYear(ItemName)
                Case Else
                    If RecordNumber = 0 Then Exit Function
                    With Me.WeeklySchedule
                        RET = .DataBody(RecordNumber, .ItemNames.Item(ItemName))
                    End With
            End Select
        Case WeekName Like "—ˆT"
            If RecordNumber = 0 Then Exit Function
            Select Case True
                Case IsDateSingleCell(ItemName, PrintingAddress)
                    RET = Me.WeeklySchedule.NextWeekFirstDate
                Case ItemName Like "T”"
                    RET = Me.WeeklySchedule.NextWeekClassWeeks
                Case TotalType Like "TŒv"
                    RET = Me.WeeklySchedule.NextWeekTotalClassHourOfWeek(ItemName)
                Case TotalType Like "—ÝŒv"
                    RET = Me.WeeklySchedule.NextWeekTotalClassHourOfYear(ItemName)
                Case Else
                    With Me.WeeklySchedule
                        RET = .DataBody(RecordNumber, .ItemNames.Item(ItemName))
                    End With
            End Select
    End Select
    If IsEmpty(RET) Then Exit Function
    If VisibledItemName Then RET = "y" & ItemName & "z" & RET
    ReturnWeeklyScheduleValue = RET
End Function
Private Function ReturnMonthlyScheduleValue(ByVal RecordNumber As Long, ByVal ItemName As String, ByVal VisibledItemName As Boolean) As Variant
    If RecordNumber = 0 Then Exit Function
    Dim RET As Variant
    With Me.MonthlySchedule
        RET = .DataBody(RecordNumber, .ItemNames.Item(ItemName))
    End With
    If IsEmpty(RET) Then Exit Function
    If VisibledItemName Then RET = "y" & ItemName & "z" & RET
    ReturnMonthlyScheduleValue = RET
End Function
Private Function ReturnYearlyScheduleValue(ByVal RecordNumber As Long, ByVal ItemName As String, ByVal VisibledItemName As Boolean, _
ByVal MonthName As String) As Variant
    If RecordNumber = 0 Then Exit Function
    Dim RET As Variant
    With Me.YearlySchedule
        If ItemName Like "Žö‹Æ“ú" Then
            RET = .TotalClassHoursOfMonth(MonthName)
        Else
            RET = .DataBody(RecordNumber, .ItemNames.Item(ItemName))
        End If
    End With
    If IsEmpty(RET) Then Exit Function
    If VisibledItemName Then RET = "y" & ItemName & "z" & RET
    ReturnYearlyScheduleValue = RET
End Function
Private Function AddFontSizeDownAddresses() As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim BufSheet As Worksheet
    Dim BufMaxFontSize As Long
    Dim BufMinFontSize As Long
    Set BufSheet = Me.Sheet
    BufMaxFontSize = Me.MaxFontSize
    BufMinFontSize = Me.MinFontSize
    Dim FontSizing As ClassFontSizing
    Set FontSizing = New ClassFontSizing
    Set FontSizing.Sheet = Me.Sheet
    With Me.PrintingItem
        Dim i As Long
        For i = 1 To UBound(.DataBody)
            Dim PrintingRange As Range
            Set PrintingRange = BufSheet.Range(.Itemˆóü”ÍˆÍ(i))
            If EnbledSizeDown(.Item•¶ŽšƒTƒCƒYk¬(i)) Then
                Dim Target As Range
                Dim UpperLeft As Range
                For Each Target In PrintingRange
                    Set UpperLeft = Target.MergeArea.Item(1)
                    If Target.Address Like UpperLeft.Address Then
                        If FontSizing.ShouldDown(Target.Value, Target) Then
                            Dim TargetAddress As String
                            TargetAddress = Target.Address(False, False)
                            If Not RET.Exists(TargetAddress) Then
                                Dim Text As String
                                Dim DownSize As Double
                                Text = ReturnFontSizingText(Target.Value)
                                DownSize = FontSizing.DownedFontSize(Text, Target, BufMaxFontSize, BufMinFontSize)
                                RET.Add TargetAddress, DownSize
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End With
    Set AddFontSizeDownAddresses = RET
End Function
Private Sub ReplaceVbLf(ByVal Addresses As Dictionary)
    On Error GoTo ThrowError
    Const ProcedureName As String = "ReplaceVbLf"
    Dim DataBody() As Variant
    DataBody = Me.UsedPrintingRange.Value
    Dim BufSheet As Worksheet
    Dim BufPositions As Dictionary
    Set BufSheet = Me.Sheet
    Set BufPositions = Me.ArrayPositions
    Dim Target As Variant
    For Each Target In Addresses
        Dim R As Long
        Dim C As Long
        R = BufSheet.Range(BufPositions.Item(Target)).Row
        C = BufSheet.Range(BufPositions.Item(Target)).Column
        DataBody(R, C) = LIB.ReplaceVbLf¨LineBreakChar(DataBody(R, C), "^")
    Next
    Me.UsedPrintingRange.Value = DataBody
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Function ToVerticalWritingDate(ByVal Target As Date) As String
    Dim Letters(7) As String
    Letters(0) = Month(Target)
    Letters(1) = "ŒŽ"
    Letters(2) = Day(Target)
    Letters(3) = "“ú"
    Letters(4) = vbNullString
    Letters(5) = WeekdayName(Weekday(Target), True)
    Letters(6) = "—j"
    Letters(7) = "“ú"
    ToVerticalWritingDate = Join(Letters, vbLf)
End Function
Private Function ReturnFontSizingText(ByVal Text As Variant) As String
    Dim RET As String
    If EnabledReplaceVbLf Then
        RET = LIB.ReplaceVbLf¨LineBreakChar(Text, "^")
    Else
        RET = CStr(Text)
    End If
    ReturnFontSizingText = RET
End Function
Private Function IsSingleCell(ByVal Target As Range) As Boolean
    If Target.Count > 1 Then Exit Function
    IsSingleCell = True
End Function
Private Function IsDateSingleCell(ByVal ItemName As String, ByVal CellAddress As String) As Boolean
    If Not ItemName Like "“ú•t" Then Exit Function
    If InStr(CellAddress, ":") > 0 Then Exit Function
    IsDateSingleCell = True
End Function
Private Function HasMergeCell(ByVal Target As Range) As Boolean
    If Target.Count = 1 Then Exit Function
    Dim BUF As Range
    Dim UpperLeft As Range
    For Each BUF In Target
        Set UpperLeft = BUF.MergeArea.Item(1)
        If BUF.Address Like UpperLeft.Address Then
            If Not BUF.MergeCells Then Exit Function
        End If
    Next
    HasMergeCell = True
End Function
Private Function IsContinuousSingleCell(ByVal Target As Range) As Boolean
    If Target.Count = 1 Then Exit Function
    Dim BUF As Range
    For Each BUF In Target
        If BUF.MergeCells Then Exit Function
    Next
    IsContinuousSingleCell = True
End Function
Private Function IsDateContinuousCell(ByVal PrintingRange As Range, ByVal ItemName As String) As Boolean
    If Not ItemName Like "“ú•t" Then Exit Function
    If PrintingRange.Count = 1 Then Exit Function
    IsDateContinuousCell = True
End Function
Private Function IsValidDate(ByVal Target As Variant) As Boolean
    If IsEmpty(Target) Then Exit Function
    If IsDate(Target) Or IsNumeric(Target) Then IsValidDate = True
End Function
Private Function EnabledReplaceVbLf() As Boolean
    If Me.ReplacedVbLf Like "‚·‚é" Then EnabledReplaceVbLf = True
End Function
Private Function EnbledSizeDown(ByVal StateName As String) As Boolean
    If Not StateName Like "‚·‚é" Then Exit Function
    EnbledSizeDown = True
End Function


