VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Leave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Form")
Option Explicit
Implements ILeave
Private ListViewControl As IListViewControl
Private ListBoxDayDutyOrder As MSForms.ListBox
Private OptionButton“Á‹x As MSForms.OptionButton
Private OptionButtonˆç‹x As MSForms.OptionButton
Private OptionButton•a‹x As MSForms.OptionButton
Private DataBody() As Variant
Private TargetName As String
Private DisplayName As String
Private LeaveTypeName As String
Private FirstDate As Date
Private LastDate As Date
Private Sub Class_Initialize()
    With FormExBB
        Set ListViewControl = New ListViewControl
        Set ListViewControl.TargetListView = .Controls("ListView’·Šú‹x‰É")
        Set ListBoxDayDutyOrder = .Controls("ListBox“ú’¼‡")
        Set OptionButton“Á‹x = .Controls("OptionButton“Á‹x")
        Set OptionButtonˆç‹x = .Controls("OptionButtonˆç‹x")
        Set OptionButton•a‹x = .Controls("OptionButton•a‹x")
    End With
    With ListViewControl
        TargetName = CStr(.GetItem("æ“¾Ò"))
        LeaveTypeName = CStr(.GetItem("í•Ê"))
        FirstDate = CDate(.GetItem("ŠJn“ú"))
        LastDate = CDate(.GetItem("I—¹“ú"))
        DisplayName = CStr(.GetItem("•\¦–¼"))
        DataBody = .GetItems
    End With
End Sub
Private Sub ILeave_Register(ByVal BasisDate As Date, ByVal MsgTitle As String)
    TargetName = LIB.SelectedItemJustOne(ListBoxDayDutyOrder)
    FirstDate = LIB.ReturnDateByInputBox("ŠJn“ú‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B", MsgTitle, BasisDate)
    If FirstDate = 0 Then Exit Sub
    LastDate = LIB.ReturnDateByInputBox("I—¹“ú‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B", MsgTitle, BasisDate)
    If LastDate = 0 Then Exit Sub
    If LIB.InvalidDateFirstToLast(BasisDate, FirstDate, LastDate, MsgTitle) Then Exit Sub
    DisplayName = InputBox("“ú“™‚Éˆóš‚·‚é•\¦–¼‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B", MsgTitle, TargetName)
    If DisplayName Like vbNullString Then Exit Sub
    LeaveTypeName = GetLeaveTypeName
    Dim RecordList As Dictionary
    Set RecordList = New Dictionary
    With RecordList
        .Add "æ“¾Ò", TargetName
        .Add "í•Ê", LeaveTypeName
        .Add "ŠJn“ú", FirstDate
        .Add "I—¹“ú", LastDate
        .Add "•\¦–¼", DisplayName
    End With
    With ListViewControl
        .AddItem RecordList
        DataBody = .GetItems
    End With
End Sub
Private Sub ILeave_Remove()
    ListViewControl.RemoveItem
End Sub
Private Function ILeave_DataBody() As Variant()
    ILeave_DataBody = DataBody
End Function
Private Function ILeave_TargetName() As String
    ILeave_TargetName = TargetName
End Function
Private Function ILeave_DisplayName() As String
    ILeave_DisplayName = DisplayName
End Function
Private Function ILeave_LeaveTypeName() As String
    ILeave_LeaveTypeName = LeaveTypeName
End Function
Private Function ILeave_FirstDate() As Date
    ILeave_FirstDate = FirstDate
End Function
Private Function ILeave_LastDate() As Date
    ILeave_LastDate = LastDate
End Function
Private Function GetLeaveTypeName() As String
    Dim RET As String
    If CBool(OptionButton“Á‹x.Value) Then
        RET = OptionButton“Á‹x.Caption
    ElseIf CBool(OptionButtonˆç‹x.Value) Then
        RET = OptionButtonˆç‹x.Caption
    ElseIf CBool(OptionButton•a‹x.Value) Then
        RET = OptionButton•a‹x.Caption
    End If
    GetLeaveTypeName = RET
End Function
