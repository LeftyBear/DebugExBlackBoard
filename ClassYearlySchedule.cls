VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassYearlySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("Schdule")
Option Explicit
Private BasisDate_      As Date
Private RecordNumbers_  As Collection
Private ItemNames_      As Dictionary
Private DataBody_       As Variant
Private Const MAX_DAYS  As Long = 31
Public Property Let BasisDate(ByVal Target As Date)
    BasisDate_ = Target
End Property
Public Property Get BasisDate() As Date
    BasisDate = BasisDate_
End Property
Public Function RecordNumbers() As Collection
    Set RecordNumbers = RecordNumbers_
End Function
Public Sub GetReadyRecordNumbers()
    Set RecordNumbers_ = AddRecordNumbers
End Sub
Public Property Get ItemNames() As Dictionary
    Set ItemNames = ItemNames_
End Property
Public Property Set ItemNames(ByVal Target As Dictionary)
    Set ItemNames_ = Target
End Property
Public Property Get DataBody() As Variant
    DataBody = DataBody_
End Property
Public Property Let DataBody(ByVal Target As Variant)
    DataBody_ = Target
End Property
Public Function TotalClassHoursOfMonth(ByVal MonthName As String) As Long
    Dim FirstDate As Date
    Dim LastDate As Date
    FirstDate = DateSerial(TargetYear(CLng(MonthName) - 3), CLng(MonthName), 1)
    LastDate = LIB.Dateç°åéññì˙(FirstDate)
    Dim FirstRecordNumber As Long
    Dim LastRecordNumber As Long
    FirstRecordNumber = LIB.DateDiff4åé1ì˙(FirstDate)
    LastRecordNumber = LIB.DateDiff4åé1ì˙(LastDate)
    With Schedule
        Dim FirstRange As Range
        Dim LastRange As Range
        Set FirstRange = .TableRecordItem(FirstRecordNumber, "éˆã∆ì˙")
        Set LastRange = .TableRecordItem(LastRecordNumber, "éˆã∆ì˙")
        If FirstRange Is Nothing Or LastRange Is Nothing Then Exit Function
        Dim Target As Range
        Set Target = .Range(FirstRange, LastRange)
    End With
    TotalClassHoursOfMonth = WorksheetFunction.CountIf(Target, True)
End Function
Private Function AddRecordNumbers() As Collection
    Dim RET As Collection
    Set RET = New Collection
    Dim RecordNumber As Long
    RecordNumber = 1
    Dim Months As Long
    For Months = 1 To 12
        Dim Days As Long
        For Days = 1 To MAX_DAYS
            Dim TargetDate As Date
            Dim LastRecordNumber As Long
            TargetDate = DateSerial(TargetYear(Months), TargetMonth(Months), 1)
            LastRecordNumber = LIB.DateDiff4åé1ì˙(LIB.Dateç°åéññì˙(TargetDate))
            If RecordNumber <= LastRecordNumber Then
                RET.Add RecordNumber
                RecordNumber = RecordNumber + 1
            Else
                RET.Add 0
            End If
        Next
    Next
    Set AddRecordNumbers = RET
End Function
Private Function TargetMonth(ByVal Months As Long) As Long
    Dim RET As Long
    Select Case Months
        Case 1 To 9
            RET = Months + 3
        Case 10 To 12
            RET = Months - 9
    End Select
    TargetMonth = RET
End Function
Private Function TargetYear(ByVal Months As Long) As Long
    Dim RET As Long
    Dim WesternYear As Long
    WesternYear = Year(LIB.Date4åé1ì˙(Me.BasisDate))
    Select Case Months
        Case 1 To 9
            RET = WesternYear
        Case 10 To 12
            RET = WesternYear + 1
        Case -2 To 0
            RET = WesternYear + 1
    End Select
    TargetYear = RET
End Function

