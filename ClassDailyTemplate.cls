VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDailyTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Template")
Option Explicit
Private Book_               As Workbook
Private TargetSheet_        As Worksheet
Private UsedPrintingRange_  As Range
Private PrintingItem_       As ClassPrintingItem
Private DailySchedule_      As ClassDailySchedule
Private HasPrintingItem_    As Boolean
Private HasSheet_           As Boolean
Private Const MSG_TITLE     As String = "日誌印刷"
Private Const MAX_FONTSIZE  As Long = 11
Private Const MIN_FONTSIZE  As Long = 8
Private Const REPLACED_VbLf As String = "しない"
Public Property Get Book() As Workbook
    Set Book = Book_
End Property
Public Property Set Book(ByVal Target As Workbook)
    Set Book_ = Target
End Property
Public Property Get TargetSheet() As Worksheet
    Set TargetSheet = TargetSheet_
End Property
Public Property Set TargetSheet(ByVal Target As Worksheet)
    Set TargetSheet_ = Target
End Property
Public Property Get UsedPrintingRange() As Range
    Set UsedPrintingRange = UsedPrintingRange_
End Property
Public Property Set UsedPrintingRange(ByVal Target As Range)
    Set UsedPrintingRange_ = Target
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
Public Property Get HasPrintingItem() As Boolean
    HasPrintingItem = HasPrintingItem_
End Property
Public Property Let HasPrintingItem(ByVal State As Boolean)
    HasPrintingItem_ = State
End Property
Public Property Get HasSheet() As Boolean
    HasSheet = HasSheet_
End Property
Public Property Let HasSheet(ByVal State As Boolean)
    HasSheet_ = State
End Property
Private Sub Class_Initialize()
    On Error Resume Next
    Dim Directory As IDirectory
    Set Directory = New Dictionary
    Set Me.Book = Workbooks.Open(FileName:=Directory.DailyTemplateFile, ReadOnly:=True)
    Set Me.TargetSheet = Me.Book.Worksheets.Item(1)
    Set Me.PrintingItem = New ClassPrintingItem
    Me.PrintingItem.ReadFile Directory.DailyTemplateSettingFile
    On Error GoTo 0
    With New ClassTemplateLinkLibrary
        Me.HasSheet = .HasSheet(Me.TargetSheet, MSG_TITLE)
        Me.HasPrintingItem = .HasPrintingItem(Me.PrintingItem, MSG_TITLE)
        If Not Me.HasSheet Then Exit Sub
        If Not Me.HasPrintingItem Then Exit Sub
        Set .TargetSheet = Me.TargetSheet
        Set .PrintingItem = Me.PrintingItem
        Set Me.UsedPrintingRange = .AddUsedPrintingRange
    End With
End Sub
Public Sub SetUpTemplate(ByVal DailySchedule As ClassDailySchedule, ByVal LinkLibrary As ClassTemplateLinkLibrary)
    Set Me.DailySchedule = DailySchedule
    With LinkLibrary
        Set .DailySchedule = Me.DailySchedule
        Set .TargetSheet = Me.TargetSheet
        Set .PrintingItem = Me.PrintingItem
        Set .UsedPrintingRange = Me.UsedPrintingRange
    End With
End Sub
Public Function DataBody(ByVal LinkLibrary As ClassTemplateLinkLibrary) As Variant()
    Dim RecordNumbers As Collection
    Set RecordNumbers = Me.DailySchedule.RecordNumbers
    With LinkLibrary
        Set .RecordNumbers = .AddRecordNumbers(RecordNumbers)
        Set .PrintingItemIndexes = .AddPrintingItemIndexes
        Set .ArrayPositions = .AddArrayPositions
        Set .ScheduleValues = .AddScheduleValues(Daily)
        DataBody = .AddDataBody
    End With
End Function
Public Sub DoFormat(ByVal LinkLibrary As ClassTemplateLinkLibrary)
    With LinkLibrary
        .MaxFontSize = MAX_FONTSIZE
        .MinFontSize = MIN_FONTSIZE
        .ReplacedVbLf = REPLACED_VbLf
        .Color日付単一セルby日直日
        .Size文字サイズ縮小
    End With
End Sub
