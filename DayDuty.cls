VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DayDuty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sheet")
Option Explicit
Implements IDayDuty
Private FirstDate       As Date
Private LastDate_       As Date
Private RecordList      As Dictionary
Private Order()         As Variant
Private TargetName      As String
Private CanInputDate    As Boolean
Private OrderNumber_    As Long
Private Property Let IDayDuty_FirstDate(ByVal RHS As Date)
    FirstDate = RHS
End Property
Private Property Get IDayDuty_FirstDate() As Date
    IDayDuty_FirstDate = FirstDate
End Property
Private Property Set IDayDuty_RecordList(ByVal RHS As Scripting.IDictionary)
    Set RecordList = RHS
End Property
Private Property Let IDayDuty_Order(ByRef RHS() As Variant)
    Order = RHS
End Property
Private Property Let IDayDuty_TargetName(ByVal RHS As String)
    TargetName = RHS
End Property
Private Property Let IDayDuty_CanInputDate(ByVal RHS As Boolean)
    CanInputDate = RHS
End Property
Private Function IDayDuty_DataBody() As Variant()
    Const MsgTitle As String = "日直割当"
    If Not IsValidDateAndTargetName(MsgTitle) Then Exit Function
    If Not HasAllowed(MsgTitle) Then Exit Function
    ResetOrderNumber
    Dim RET() As Variant
    ReDim RET(1 To LastDate_ - IDayDuty_FirstDate + 1, 1 To 1)
    Dim R As Long
    R = 1
    Dim Target As Date
    For Target = IDayDuty_FirstDate To LastDate_
        If CBool(RecordList.Exists(Target)) Then
            RET(R, 1) = Rotation
        End If
        R = R + 1
    Next
    IDayDuty_DataBody = RET
End Function
Private Sub ResetOrderNumber()
    OrderNumber_ = 0
    Dim i As Long
    For i = LBound(Order) To UBound(Order)
        If TargetName = Rotation Then Exit For
    Next
    OrderNumber_ = OrderNumber_ - 1
    If UBound(Order) < LBound(Order) + OrderNumber_ Then OrderNumber_ = 0
End Sub
Private Function Rotation() As String
    Rotation = Order(LBound(Order) + OrderNumber_)
    OrderNumber_ = OrderNumber_ + 1
    If UBound(Order) < LBound(Order) + OrderNumber_ Then OrderNumber_ = 0
End Function
Private Function PreviousName() As String
    Dim RET As String
    If 1 < LIB.DateDiff4月1日(IDayDuty_FirstDate) Then
        Dim Apr1st As Date
        Apr1st = LIB.Date4月1日(IDayDuty_FirstDate)
        Dim Target As Date
        For Target = IDayDuty_FirstDate - 1 To Apr1st Step -1
            If RecordList.Exists(Target) Then
                If Not RecordList.Item(Target) Like vbNullString Then
                    RET = RecordList.Item(Target)
                    Exit For
                End If
            End If
        Next
    Else
        RET = Order(UBound(Order))
    End If
    PreviousName = RET
End Function
Private Function NextName() As String
    OrderNumber_ = 0
    Dim i As Long
    For i = LBound(Order) To UBound(Order)
        If TargetName = Rotation Then Exit For
    Next
    If UBound(Order) < i + 1 Then
        Dim NextNumber As Long
        NextNumber = LBound(Order)
    Else
        NextNumber = i + 1
    End If
    NextName = Order(NextNumber)
End Function
Private Function IsValidDateAndTargetName(ByVal MsgTitle As String) As Boolean
    If CanInputDate Then
        IDayDuty_FirstDate = LIB.ReturnDateByInputBox("開始日を入力してください。", MsgTitle, IDayDuty_FirstDate)
        If IDayDuty_FirstDate = 0 Then Exit Function
        LastDate_ = LIB.ReturnDateByInputBox("終了日を入力してください。", MsgTitle, IDayDuty_FirstDate)
        If LastDate_ = 0 Then Exit Function
        If Not IsValidDate(MsgTitle) Then Exit Function
    Else
        IDayDuty_FirstDate = LIB.Date今月1日(IDayDuty_FirstDate)
        LastDate_ = LIB.Date今月末日(IDayDuty_FirstDate)
        TargetName = PreviousName
        TargetName = NextName
    End If
    IsValidDateAndTargetName = True
End Function
Private Function IsValidDate(ByVal MsgTitle As String) As Boolean
    Dim Apr1st As Date
    Apr1st = LIB.Date4月1日(IDayDuty_FirstDate)
    Dim Mar31th As Date
    Mar31th = LIB.Date3月31日(IDayDuty_FirstDate)
    If Apr1st <= IDayDuty_FirstDate And IDayDuty_FirstDate <= Mar31th Then
        If Apr1st <= LastDate_ And LastDate_ <= Mar31th Then
            If IDayDuty_FirstDate <= LastDate_ Then IsValidDate = True
        End If
    End If
    MsgBox "開始日または終了日、もしくは両方に誤りがあります。", vbExclamation, MsgTitle
End Function
Private Function HasAllowed(ByVal MsgTitle As String) As Boolean
    Dim RES As Long
    RES = MsgBox(TargetName & " から順に" & vbCrLf & IDayDuty_FirstDate & "から" & LastDate_ & "まで" & _
        vbCrLf & "日直を割り当てますか？" & vbCrLf & "（すでに割り当てられている日直は上書きします。）", _
        vbYesNo + vbQuestion, MsgTitle)
    If RES = vbYes Then HasAllowed = True
End Function

