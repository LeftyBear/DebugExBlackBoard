VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDayDuty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
Private DayDutyNames_       As Dictionary
Private TargetList_         As Variant
Private TargetName_         As String
Private BasisDate_          As Date
Private FirstDate_          As Date
Private LastDate_           As Date
Private OrderNumber_        As Long
Public Property Set DayDutyNames(ByVal Target As Dictionary)
    Set DayDutyNames_ = Target
End Property
Public Property Get DayDutyNames() As Dictionary
    Set DayDutyNames = DayDutyNames_
End Property
Public Property Let TargetList(ByVal Target As Variant)
    TargetList_ = Target
End Property
Public Property Get TargetList() As Variant
    TargetList = TargetList_
End Property
Public Property Let TargetName(ByVal Target As String)
    TargetName_ = Target
End Property
Public Property Get TargetName() As String
    TargetName = TargetName_
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
Public Sub AddTargetName()
    Me.TargetName = PreviousName
    Me.TargetName = NextName
End Sub
Public Function AddDayDutys() As Variant()
    Dim BufDayDutys As Dictionary
    Set BufDayDutys = Me.DayDutyNames
    NewOrderNumber
    Dim RET() As Variant
    ReDim RET(1 To Me.LastDate - Me.FirstDate + 1, 1 To 1)
    Dim i As Long
    i = 1
    Dim Dates As Date
    For Dates = Me.FirstDate To Me.LastDate
        If CBool(BufDayDutys.Exists(Dates)) Then
            RET(i, 1) = Rotation
        End If
        i = i + 1
    Next
    AddDayDutys = RET
End Function
Private Function PreviousName() As String
    Dim RET As String
    Dim BufDayDutyNames As Dictionary
    Set BufDayDutyNames = Me.DayDutyNames
    If LIB.DateDiff4ŒŽ1“ú(Me.FirstDate) > 1 Then
        Dim Dates As Date
        For Dates = Me.FirstDate - 1 To LIB.Date4ŒŽ1“ú(Me.FirstDate) Step -1
            If BufDayDutyNames.Exists(Dates) Then
                If Not BufDayDutyNames.Item(Dates) Like vbNullString Then
                    RET = BufDayDutyNames.Item(Dates)
                    Exit For
                End If
            End If
        Next
    End If
    If RET Like vbNullString Then RET = Me.TargetList(UBound(Me.TargetList))
    PreviousName = RET
End Function
Private Function NextName() As String
    OrderNumber_ = 0
    Dim i As Long
    For i = LBound(Me.TargetList) To UBound(Me.TargetList)
        If Me.TargetName = Rotation Then Exit For
    Next
    Dim IndexNumber As Long
    '@Ignore IIfSideEffect
    IndexNumber = IIf(i + 1 > UBound(Me.TargetList), LBound(Me.TargetList), i + 1)
    NextName = Me.TargetList(IndexNumber)
End Function
Private Sub NewOrderNumber()
    OrderNumber_ = 0
    Dim i As Long
    For i = LBound(Me.TargetList) To UBound(Me.TargetList)
        If Me.TargetName = Rotation Then Exit For
    Next
    OrderNumber_ = OrderNumber_ - 1
    If LBound(Me.TargetList) + OrderNumber_ > UBound(Me.TargetList) Then OrderNumber_ = 0
End Sub
Private Function Rotation() As String
    Rotation = Me.TargetList(LBound(Me.TargetList) + OrderNumber_)
    OrderNumber_ = OrderNumber_ + 1
    If LBound(Me.TargetList) + OrderNumber_ > UBound(Me.TargetList) Then OrderNumber_ = 0
End Function
Public Function IsValidDate() As Boolean
    Dim Apr1st As Date
    Dim Mar31th As Date
    Apr1st = LIB.Date4ŒŽ1“ú(Me.BasisDate)
    Mar31th = LIB.Date3ŒŽ31“ú(Me.BasisDate)
    If Apr1st <= Me.FirstDate And Me.FirstDate <= Mar31th Then
        If Apr1st <= Me.LastDate And Me.LastDate <= Mar31th Then
            If Me.FirstDate <= Me.LastDate Then IsValidDate = True
        End If
    End If
End Function

