VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassRegularSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
Private RegularDayList_     As Variant
Private RegularWeekDayList_ As Variant
Private BasisDate_          As Date
Private FirstDate_          As Date
Private LastDate_           As Date
Private Const COLUMNSójì˙   As Long = 0
Private Const COLUMNSì˙Ç…Çø As Long = 0
Private Const COLUMNSó\íË   As Long = 1
Public Property Let RegularDayList(ByVal Target As Variant)
    RegularDayList_ = Target
End Property
Public Property Get RegularDayList() As Variant
    RegularDayList = RegularDayList_
End Property
Public Property Let RegularWeekDayList(ByVal Target As Variant)
    RegularWeekDayList_ = Target
End Property
Public Property Get RegularWeekDayList() As Variant
    RegularWeekDayList = RegularWeekDayList_
End Property
Public Property Let BasisDate(ByVal Target As Date)
    BasisDate_ = Target
End Property
Public Property Get BasisDate() As Date
    BasisDate = BasisDate_
End Property
Public Property Let FirstDate(ByVal Target As Date)
    FirstDate_ = Target
End Property
Public Property Get FirstDate() As Date
    FirstDate = FirstDate_
End Property
Public Property Let LastDate(ByVal Target As Date)
    LastDate_ = Target
End Property
Public Property Get LastDate() As Date
    LastDate = LastDate_
End Property
Public Function AddRegulars() As Variant()
    Dim Events As Dictionary
    Set Events = Schedule.TableColumnItemList("çséñ")
    If Events Is Nothing Then Exit Function
    Dim RET() As Variant
    ReDim RET(1 To Me.LastDate - Me.FirstDate + 1, 1 To 1)
    Dim RegularDayDataBody() As Variant
    Dim RegularWeekDayDataBody() As Variant
    RegularDayDataBody = AddRegularDayDataBody
    RegularWeekDayDataBody = AddRegularWeekDayDataBody
    Dim i As Long
    i = 1
    Dim Dates As Date
    For Dates = Me.FirstDate To Me.LastDate
        RET(i, 1) = Events.Item(Dates)
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
Public Function IsValidDate() As Boolean
    Dim Apr1st As Date
    Dim Mar31th As Date
    Apr1st = LIB.Date4åé1ì˙(Me.BasisDate)
    Mar31th = LIB.Date3åé31ì˙(Me.BasisDate)
    If Apr1st <= Me.FirstDate And Me.FirstDate <= Mar31th Then
        If Apr1st <= Me.LastDate And Me.LastDate <= Mar31th Then
            If Me.FirstDate <= Me.LastDate Then IsValidDate = True
        End If
    End If
End Function
Private Function AddRegularDays() As Dictionary
    Dim BUF As Dictionary
    Set BUF = New Dictionary
    Dim i As Long
    For i = LBound(Me.RegularDayList) To UBound(Me.RegularDayList)
        If Me.RegularDayList(i, 0) = "åéññ" Then
            BUF.Add CLng(Day(Me.LastDate)), Me.RegularDayList(i, COLUMNSó\íË)
        Else
            BUF.Add CLng(Me.RegularDayList(i, COLUMNSì˙Ç…Çø)), Me.RegularDayList(i, COLUMNSó\íË)
        End If
    Next
    Set AddRegularDays = BUF
End Function
Private Function AddRegularWeekDays() As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim i As Long
    For i = LBound(Me.RegularWeekDayList) To UBound(Me.RegularWeekDayList)
        RET.Add LIB.ReplaceÅõójì˙Å®WeekDay(Me.RegularWeekDayList(i, COLUMNSójì˙)), Me.RegularWeekDayList(i, COLUMNSó\íË)
    Next
    Set AddRegularWeekDays = RET
End Function
Private Function AddRegularDayDataBody() As Variant()
    Dim RET() As Variant
    ReDim RET(1 To Me.LastDate - Me.FirstDate + 1, 1 To 1)
    Dim RegularDays As Dictionary
    Set RegularDays = AddRegularDays
    Dim i As Long
    i = 1
    Dim Dates As Date
    For Dates = Me.FirstDate To Me.LastDate
        If RegularDays.Exists(Day(Dates)) Then
            RET(i, 1) = RegularDays.Item(Day(Dates))
        End If
        i = i + 1
    Next
    AddRegularDayDataBody = RET
End Function
Private Function AddRegularWeekDayDataBody() As Variant()
    Dim RET() As Variant
    ReDim RET(1 To Me.LastDate - Me.FirstDate + 1, 1 To 1)
    Dim RegularWeekDays As Dictionary
    Set RegularWeekDays = AddRegularWeekDays
    Dim i As Long
    i = 1
    Dim Dates As Date
    For Dates = Me.FirstDate To Me.LastDate
        If RegularWeekDays.Exists(Weekday(Dates)) Then
            RET(i, 1) = RegularWeekDays.Item(Weekday(Dates))
        End If
        i = i + 1
    Next
    AddRegularWeekDayDataBody = RET
End Function

