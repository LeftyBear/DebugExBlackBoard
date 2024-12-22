VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sheet")
Option Explicit
Implements IClassWeek
Private Function IClassWeek_Count(ByVal TargetDate As Date, ByVal RecordList As Scripting.IDictionary) As String
    Dim FirstMonday As Date
    FirstMonday = LIB.Dateç°èTèâì˙(LIB.Date4åé1ì˙(TargetDate))
    Dim ClassWeekList As Dictionary
    Set ClassWeekList = ClassWeeksList(FirstMonday, RecordList)
    Dim TargetWeek As Long
    TargetWeek = DateDiff("w", CVar(FirstMonday), CVar(TargetDate)) + 1
    IClassWeek_Count = ClassWeekList.Item(TargetWeek)
End Function
Private Function ClassWeeksList(ByVal FirstMonday As Date, ByVal RecordList As Dictionary) As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim ClassDay As Variant
    For Each ClassDay In RecordList
        Dim HasClassDayOfThisWeek As Boolean
        If CBool(RecordList.Item(ClassDay)) Then HasClassDayOfThisWeek = True
        If Weekday(ClassDay) = vbSunday Then
            If HasClassDayOfThisWeek Then
                Dim ClassWeekCount As Long
                ClassWeekCount = ClassWeekCount + 1
                Dim ClassWeeksString As String
                ClassWeeksString = CStr(ClassWeekCount)
            Else
                ClassWeeksString = vbNullString
            End If
            RET.Add DateDiff("w", CVar(FirstMonday), ClassDay) + 1, ClassWeeksString
        End If
        HasClassDayOfThisWeek = False
    Next
    Set ClassWeeksList = RET
End Function

