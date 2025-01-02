VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegularSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Form")
Option Explicit
Implements IRegularSchedule
Private ListViewDay                     As IListViewControl
Private ListViewWeekday                 As IListViewControl
Private OptionButtoníËä˙ó\íËì˙ïtéwíË    As MSForms.OptionButton
Private DayAndWeekdayList               As Dictionary
Private FirstDate                       As Date
Private LastDate                        As Date
Private DayBody()                       As Variant
Private WeekdayBody()                   As Variant
Private Sub Class_Initialize()
    With FormExBB
        Set ListViewDay = New ListViewControl
        Set ListViewDay.TargetListView = .Controls("ListViewíËä˙ó\íËì˙Ç…Çø")
        Set ListViewWeekday = New ListViewControl
        Set ListViewWeekday.TargetListView = .Controls("ListViewíËä˙ó\íËójì˙")
        Set OptionButtoníËä˙ó\íËì˙ïtéwíË = .Controls("OptionButtoníËä˙ó\íËì˙ïtéwíË")
    End With
    DayBody = ListViewDay.GetItems
    WeekdayBody = ListViewWeekday.GetItems
End Sub
Private Sub IRegularSchedule_Register(ByVal BasisDate As Date, ByVal MsgTitle As String)
    If LIB.IsNoneArray(DayBody) And LIB.IsNoneArray(WeekdayBody) Then MsgBox "îΩâfÇ∑ÇÈíËä˙ó\íËÇÕìoò^Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB", vbInformation, MsgTitle: Exit Sub
    If CBool(OptionButtoníËä˙ó\íËì˙ïtéwíË.Value) Then
        FirstDate = LIB.ReturnDateByInputBox("äJénì˙Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", MsgTitle, BasisDate)
        If FirstDate = 0 Then Exit Sub
        LastDate = LIB.ReturnDateByInputBox("èIóπì˙Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", MsgTitle, BasisDate)
        If LastDate = 0 Then Exit Sub
        If LIB.InvalidDateFirstToLast(BasisDate, FirstDate, LastDate, MsgTitle) Then Exit Sub
    Else
        FirstDate = LIB.Dateç°åé1ì˙(BasisDate)
        LastDate = LIB.Dateç°åéññì˙(BasisDate)
    End If
    Dim DayList As Dictionary
    Set DayList = ToDictionary(DayBody)
    Dim WeekdayList As Dictionary
    Set WeekdayList = ToDictionary(WeekdayBody)
    Set DayAndWeekdayList = LIB.MergeDictionary(DayList, WeekdayList)
End Sub
Private Function IRegularSchedule_DayAndWeekdayList() As Scripting.IDictionary
    Set IRegularSchedule_DayAndWeekdayList = DayAndWeekdayList
End Function
Private Function IRegularSchedule_FirstDate() As Date
    IRegularSchedule_FirstDate = FirstDate
End Function
Private Function IRegularSchedule_LastDate() As Date
    IRegularSchedule_LastDate = LastDate
End Function
Private Function ToDictionary(ByRef DataBody() As Variant) As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    
    Set ToDictionary = RET
End Function
Public Function AddRegulars() As Variant()
    Dim RET() As Variant
    ReDim RET(1 To LastDate - FirstDate + 1, 1 To 1)
    Dim RegularDayDataBody() As Variant
    Dim RegularWeekDayDataBody() As Variant
    RegularDayDataBody = AddRegularDayDataBody
    RegularWeekDayDataBody = AddRegularWeekDayDataBody
    Dim i As Long
    i = 1
    Dim Dates As Date
    For Dates = FirstDate To LastDate
        RET(i, 1) = DayAndWeekdayList.Item(Dates)
        If Not IsEmpty(RET(i, 1)) Then
            If Not IsEmpty(RegularDayDataBody(i, 1)) Then
                If Not RET(i, 1) Like "*" & RegularDayDataBody(i, 1) & "*" Then
                    RET(i, 1) = RegularDayDataBody(i, 1) & vbLf & RET(i, 1)
                End If
            End If
            If Not IsEmpty(RegularWeekDayDataBody(i, 1)) Then
                If Not RET(i, 1) Like "*" & RegularWeekDayDataBody(i, 1) & "*" Then
                    RET(i, 1) = RET(i, 1) & vbLf & RegularWeekDayDataBody(i, 1)
                End If
            End If
        Else
            If Not IsEmpty(RegularDayDataBody(i, 1)) Then
                RET(i, 1) = RegularDayDataBody(i, 1)
            End If
            If Not IsEmpty(RegularWeekDayDataBody(i, 1)) Then
                RET(i, 1) = RegularWeekDayDataBody(i, 1)
            End If
        End If
        i = i + 1
    Next
    AddRegulars = RET
End Function
Private Function AddRegularDays() As Dictionary
    Dim BUF As Dictionary
    Set BUF = New Dictionary
    Dim i As Long
    For i = LBound(DayBody) To UBound(DayBody)
        If DayBody(i, 0) = "åéññ" Then
            BUF.Add CLng(Day(LastDate)), DayBody(i, 2)
        Else
            BUF.Add CLng(DayBody(i, 1)), DayBody(i, 2)
        End If
    Next
    Set AddRegularDays = BUF
End Function
Private Function AddRegularWeekDays() As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim i As Long
    For i = LBound(WeekdayBody) To UBound(WeekdayBody)
        RET.Add LIB.ReplaceÅõójì˙Å®WeekDay(WeekdayBody(i, 1)), WeekdayBody(i, 2)
    Next
    Set AddRegularWeekDays = RET
End Function
Private Function AddRegularDayDataBody() As Variant()
    Dim RET() As Variant
    ReDim RET(1 To LastDate - FirstDate + 1, 1 To 1)
    Dim RegularDays As Dictionary
    Set RegularDays = AddRegularDays
    Dim i As Long
    i = 1
    Dim Dates As Date
    For Dates = FirstDate To LastDate
        If RegularDays.Exists(Day(Dates)) Then
            RET(i, 1) = RegularDays.Item(Day(Dates))
        End If
        i = i + 1
    Next
    AddRegularDayDataBody = RET
End Function
Private Function AddRegularWeekDayDataBody() As Variant()
    Dim RET() As Variant
    ReDim RET(1 To LastDate - FirstDate + 1, 1 To 1)
    Dim RegularWeekDays As Dictionary
    Set RegularWeekDays = AddRegularWeekDays
    Dim i As Long
    i = 1
    Dim Dates As Date
    For Dates = FirstDate To LastDate
        If RegularWeekDays.Exists(Weekday(Dates)) Then
            RET(i, 1) = RegularWeekDays.Item(Weekday(Dates))
        End If
        i = i + 1
    Next
    AddRegularWeekDayDataBody = RET
End Function
