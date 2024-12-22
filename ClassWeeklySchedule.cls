VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassWeeklySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("Schedule")
Option Explicit
Private Sheet               As ISheetWithTable
Private BasisDate_          As Date
Private RecordNumbers_      As Collection
Private ItemNames_          As Dictionary
Private DataBody_           As Variant
Private Const MAX_DAYS      As Long = 7 * 2
Public Property Get BasisDate() As Date
    BasisDate = BasisDate_
End Property
Public Property Let BasisDate(ByVal Target As Date)
    BasisDate_ = Target
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
Private Function AddRecordNumbers() As Collection
    Dim RET As Collection
    Set RET = New Collection
    Const PastDays As Long = 2
    Dim RecordNumber As Long
    Dim BufMaxRecordNumber As Long
    Dim Containers As Long
    RecordNumber = ThisWeekFirstRecordNumber - PastDays
    BufMaxRecordNumber = MaxRecordNumber
    Containers = 1
    Dim Days As Long
    For Days = 1 To MAX_DAYS
        If 1 <= RecordNumber And RecordNumber <= BufMaxRecordNumber Then
            Set Sheet = Schedule.Sheet
            Dim TargetWeekDay As Long
            TargetWeekDay = Weekday(Sheet.DailyRecordItem(RecordNumber, "ì˙ït").Value)
            Select Case Containers * 10 + TargetWeekDay
                Case 1 * 10 + vbSaturday, 1 * 10 + vbSunday
                    If IsValidRecNumAndì˙íºì˙(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        RecordNumber = RecordNumber + 1
                    End If
                Case 1 * 10 + vbMonday
                    If IsValidRecNum(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        TakeContainer False, Containers, RecordNumber, RET
                    End If
                Case 2 * 10 + vbSunday
                    If IsValidRecNumAndì˙íºì˙(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        RecordNumber = RecordNumber + 1
                    End If
                Case 2 * 10 + vbMonday, 2 * 10 + vbTuesday
                    If IsValidRecNum(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        TakeContainer False, Containers, RecordNumber, RET
                    End If
                Case 3 * 10 + vbMonday, 3 * 10 + vbTuesday, 3 * 10 + vbWednesday
                    If IsValidRecNum(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        TakeContainer False, Containers, RecordNumber, RET
                    End If
                Case 4 * 10 + vbTuesday, 4 * 10 + vbWednesday, 4 * 10 + vbThursday
                    If IsValidRecNum(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        TakeContainer False, Containers, RecordNumber, RET
                    End If
                Case 5 * 10 + vbWednesday, 5 * 10 + vbThursday, 5 * 10 + vbFriday
                    If IsValidRecNum(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        TakeContainer False, Containers, RecordNumber, RET
                    End If
                Case 6 * 10 + vbSaturday, 6 * 10 + vbSunday
                    If IsValidRecNumAndì˙íºì˙(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        RecordNumber = RecordNumber + 1
                    End If
                Case 6 * 10 + vbMonday, 6 * 10 + vbTuesday, 6 * 10 + vbWednesday, 6 * 10 + vbThursday
                    If IsValidRecNum(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        TakeContainer False, Containers, RecordNumber, RET
                    End If
                Case 7 * 10 + vbSunday
                    If IsValidRecNumAndì˙íºì˙(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    End If
                Case 7 * 10 + vbMonday, 7 * 10 + vbTuesday, 7 * 10 + vbWednesday, 7 * 10 + vbThursday, 7 * 10 + vbFriday
                    If IsValidRecNum(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        TakeContainer False, Containers, RecordNumber, RET
                    End If
                Case Else
                    If IsValidRecNum(RecordNumber, BufMaxRecordNumber) Then
                        TakeContainer True, Containers, RecordNumber, RET
                    Else
                        TakeContainer False, Containers, RecordNumber, RET
                    End If
            End Select
        ElseIf RecordNumber < 1 Then
            Dim PastDaysCount As Long
            If PastDaysCount < PastDays Then
                RecordNumber = RecordNumber + 1
                PastDaysCount = PastDaysCount + 1
            Else
                TakeContainer False, Containers, RecordNumber, RET
            End If
        ElseIf BufMaxRecordNumber < RecordNumber Then
            TakeContainer False, Containers, RecordNumber, RET
        End If
    Next
    Set AddRecordNumbers = RET
End Function
Public Function ThisWeekTotalClassHourOfWeek(ByVal ItemName As String) As Long
    Dim FirstRecordNumer As Long
    Dim LastRecordNumber As Long
    FirstRecordNumer = ThisWeekFirstRecordNumber
    LastRecordNumber = ThisWeekLastRecordNumber
    If FirstRecordNumer < 1 Then FirstRecordNumer = 1
    If LastRecordNumber > MaxRecordNumber Then LastRecordNumber = MaxRecordNumber
    Set Sheet = Schedule.Sheet
    With Sheet
        Dim FirstRange As Range
        Dim LastRange As Range
        Set FirstRange = .DailyRecordItem(FirstRecordNumer, ItemName)
        Set LastRange = .DailyRecordItem(LastRecordNumber, ItemName)
        If FirstRange Is Nothing Or LastRange Is Nothing Then Exit Function
        Dim Target As Range
        Set Target = Schedule.Range(FirstRange, LastRange)
    End With
    ThisWeekTotalClassHourOfWeek = WorksheetFunction.Sum(Target)
End Function
Public Function NextWeekTotalClassHourOfWeek(ByVal ItemName As String) As Long
    Dim FirstRecordNumer As Long
    FirstRecordNumer = NextWeekFirstRecordNumber
    Dim LastRecordNumber As Long
    LastRecordNumber = NextWeekLastRecordNumber
    If FirstRecordNumer < 1 Then FirstRecordNumer = 1
    If LastRecordNumber > MaxRecordNumber Then LastRecordNumber = MaxRecordNumber
    Set Sheet = Schedule.Sheet
    With Sheet
        Dim FirstRange As Range
        Dim LastRange As Range
        Set FirstRange = .DailyRecordItem(FirstRecordNumer, ItemName)
        Set LastRange = .DailyRecordItem(LastRecordNumber, ItemName)
        If FirstRange Is Nothing Or LastRange Is Nothing Then Exit Function
        Dim Target As Range
        Set Target = Schedule.Range(FirstRange, LastRange)
    End With
    NextWeekTotalClassHourOfWeek = WorksheetFunction.Sum(Target)
End Function
Public Function ThisWeekTotalClassHourOfYear(ByVal ItemName As String) As Long
    Const FirstRecordNumer As Long = 1
    Dim LastRecordNumber As Long
    LastRecordNumber = ThisWeekLastRecordNumber
    If LastRecordNumber > MaxRecordNumber Then LastRecordNumber = MaxRecordNumber
    Set Sheet = Schedule.Sheet
    With Sheet
        Dim FirstRange As Range
        Dim LastRange As Range
        Set FirstRange = .DailyRecordItem(FirstRecordNumer, ItemName)
        Set LastRange = .DailyRecordItem(LastRecordNumber, ItemName)
        If FirstRange Is Nothing Or LastRange Is Nothing Then Exit Function
        Dim Target As Range
        Set Target = Schedule.Range(FirstRange, LastRange)
    End With
    ThisWeekTotalClassHourOfYear = WorksheetFunction.Sum(Target)
End Function
Public Function NextWeekTotalClassHourOfYear(ByVal ItemName As String) As Long
    Const FirstRecordNumer As Long = 1
    Dim LastRecordNumber As Long
    LastRecordNumber = NextWeekLastRecordNumber
    If LastRecordNumber > MaxRecordNumber Then LastRecordNumber = MaxRecordNumber
    Set Sheet = Schedule.Sheet
    With Sheet
        Dim FirstRange As Range
        Dim LastRange As Range
        Set FirstRange = .DailyRecordItem(FirstRecordNumer, ItemName)
        Set LastRange = .DailyRecordItem(LastRecordNumber, ItemName)
        If FirstRange Is Nothing Or LastRange Is Nothing Then Exit Function
        Dim Target As Range
        Set Target = Schedule.Range(FirstRange, LastRange)
    End With
    NextWeekTotalClassHourOfYear = WorksheetFunction.Sum(Target)
End Function
Public Function ThisWeekFirstDate() As Date
    ThisWeekFirstDate = LIB.Dateç°èTèâì˙(Me.BasisDate, vbMonday)
End Function
Public Function NextWeekFirstDate() As Date
    NextWeekFirstDate = LIB.Dateç°èTèâì˙(Me.BasisDate, vbMonday) + 7
End Function
Private Function ThisWeekLastDate() As Date
    ThisWeekLastDate = LIB.Dateç°èTññì˙(Me.BasisDate, vbSunday)
End Function
Private Function NextWeekLastDate() As Date
    NextWeekLastDate = LIB.Dateç°èTññì˙(Me.BasisDate, vbSunday) + 7
End Function
Public Function ThisWeekFirstRecordNumber() As Long
    ThisWeekFirstRecordNumber = ThisWeekFirstDate - LIB.Date4åé1ì˙(Me.BasisDate) + 1
End Function
Public Function NextWeekFirstRecordNumber() As Long
    NextWeekFirstRecordNumber = NextWeekFirstDate - LIB.Date4åé1ì˙(Me.BasisDate) + 1
End Function
Public Function ThisWeekLastRecordNumber() As Long
    ThisWeekLastRecordNumber = ThisWeekLastDate - LIB.Date4åé1ì˙(Me.BasisDate) + 1
End Function
Public Function NextWeekLastRecordNumber() As Long
    NextWeekLastRecordNumber = NextWeekLastDate - LIB.Date4åé1ì˙(Me.BasisDate) + 1
End Function
Public Function ThisWeekClassWeeks() As String
    Dim RecordList As Dictionary
    Set RecordList = Schedule.Sheet.YearlyRecordItemList("éˆã∆ì˙")
    ThisWeekClassWeeks = Schedule.ClassWeek.Count(ThisWeekFirstDate, RecordList)
End Function
Public Function NextWeekClassWeeks() As String
    Dim RecordList As Dictionary
    Set RecordList = Schedule.Sheet.YearlyRecordItemList("éˆã∆ì˙")
    Dim TargetDate As Date
    TargetDate = NextWeekFirstDate
    If LIB.Date3åé31ì˙(Me.BasisDate) < TargetDate Then TargetDate = LIB.Date3åé31ì˙(Me.BasisDate)
    NextWeekClassWeeks = Schedule.ClassWeek.Count(TargetDate, RecordList)
End Function
Private Function MaxRecordNumber() As Long
    MaxRecordNumber = LIB.Date3åé31ì˙(Me.BasisDate) - LIB.Date4åé1ì˙(Me.BasisDate) + 1
End Function
Private Sub TakeContainer(ByVal IsValid As Boolean, ByRef Containers As Long, _
ByRef RecordNumber As Long, ByVal RecordNumbers As Collection)
    If IsValid Then
        RecordNumbers.Add RecordNumber
    Else
        RecordNumbers.Add 0
    End If
    RecordNumber = RecordNumber + 1
    Containers = Containers + 1
End Sub
Private Function IsValidRecNumAndì˙íºì˙(ByVal RecordNumber As Long, ByVal MaxRecordNumber As Long) As Boolean
    Set Sheet = Schedule.Sheet
    If RecordNumber < 1 Or MaxRecordNumber < RecordNumber Then Exit Function
    If Not CBool(Sheet.DailyRecordItem(RecordNumber, "ì˙íºì˙").Value) Then Exit Function
    IsValidRecNumAndì˙íºì˙ = True
End Function
Private Function IsValidRecNum(ByVal RecordNumber As Long, ByVal MaxRecordNumber As Long) As Boolean
    If RecordNumber < 1 Or MaxRecordNumber < RecordNumber Then Exit Function
    IsValidRecNum = True
End Function


