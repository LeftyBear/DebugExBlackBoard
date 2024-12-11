VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScheduleTableCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IScheduleTableCalculator
Private Function IScheduleTableCalculator_CountedClassWeeks(ByVal Table As ITable, ByVal TargetDate As Date) As String
    Dim FirstMonday As Date
    FirstMonday = LIB.Dateç°èTèâì˙(LIB.Date4åé1ì˙(TargetDate))
    Dim ClassWeeks As Dictionary
    Set ClassWeeks = ClassWeeksList(Table, FirstMonday)
    Dim TargetWeek As Long
    TargetWeek = DateDiff("w", CVar(FirstMonday), CVar(TargetDate)) + 1
    IScheduleTableCalculator_CountedClassWeeks = ClassWeeks.Item(TargetWeek)
End Function
Private Function ClassWeeksList(ByVal Table As ITable, ByVal FirstMonday As Date) As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    With Table
        Dim Dates() As Variant
        Dates = .ColumnItems("ì˙ït")
        Dim ClassDays() As Variant
        ClassDays = .ColumnItems("éˆã∆ì˙")
    End With
    Dim i As Long
    For i = LBound(Dates, 1) To UBound(Dates, 1)
        Dim ThisWeekClassDayExists As Boolean
        If CBool(ClassDays(i, 1)) Then ThisWeekClassDayExists = True
        If Weekday(Dates(i, 1)) = vbSunday Then
            If ThisWeekClassDayExists Then
                Dim ClassWeeksCount As Long
                ClassWeeksCount = ClassWeeksCount + 1
                Dim ClassWeeksString As String
                ClassWeeksString = CStr(ClassWeeksCount)
            Else
                ClassWeeksString = vbNullString
            End If
            RET.Add DateDiff("w", CVar(FirstMonday), Dates(i, 1)) + 1, ClassWeeksString
        End If
        ThisWeekClassDayExists = False
    Next
    Set ClassWeeksList = RET
End Function

