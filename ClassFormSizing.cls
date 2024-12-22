VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFormSizing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal Hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal Hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As LongPtr
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal Hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal Hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
Private Const MODULE_NAME       As String = "ClassFormSizing"
Private TargetForm_             As FormExBB
Private EnabledMinimization_    As Boolean
Private EnabledMaximization_    As Boolean
Private EnabledResizing_        As Boolean
Private EnabledClosingButton_   As Boolean
Private TurnOffResizeEvent      As Boolean
Private Const RATE_VERTICAL     As Long = 9
Private Const RATE_HORIZONTAL   As Long = 17
Public Property Set TargetForm(ByVal Target As FormExBB)
    Set TargetForm_ = Target
End Property
Public Property Get TargetForm() As FormExBB
    Set TargetForm = TargetForm_
End Property
Public Property Get EnabledMinimization() As Boolean
    EnabledMinimization = EnabledMinimization_
End Property
Public Property Let EnabledMinimization(ByVal IsEnabled As Boolean)
    EnabledMinimization_ = IsEnabled
End Property
Public Property Get EnabledMaximization() As Boolean
    EnabledMaximization = EnabledMaximization_
End Property
Public Property Let EnabledMaximization(ByVal IsEnabled As Boolean)
    EnabledMaximization_ = IsEnabled
End Property
Public Property Get EnabledResizing() As Boolean
    EnabledResizing = EnabledResizing_
End Property
Public Property Let EnabledResizing(ByVal IsEnabled As Boolean)
    EnabledResizing_ = IsEnabled
End Property
Public Property Get EnabledClosingButton() As Boolean
    EnabledClosingButton = EnabledClosingButton_
End Property
Public Property Let EnabledClosingButton(ByVal IsEnabled As Boolean)
    EnabledClosingButton_ = IsEnabled
End Property
Public Sub DrawFormMenuBar()
    Const ProcedureName As String = "DrawFormMenuBar"
    On Error GoTo ThrowError
    Dim WindowHandle As LongPtr
    WindowHandle = FindWindow("ThunderDFrame", Me.TargetForm.Caption)
    Dim MenuStyle As LongPtr
    MenuStyle = GetWindowLongPtr(WindowHandle, (-16))
    If EnabledMinimization Then MenuStyle = MenuStyle Or &H20000
    If EnabledMaximization Then MenuStyle = MenuStyle Or &H10000
    If EnabledResizing Then MenuStyle = MenuStyle Or &H40000
    SetWindowLongPtr WindowHandle, (-16), MenuStyle
    If Not EnabledClosingButton Then
        Dim ClosingButtonStyle As LongPtr
        ClosingButtonStyle = GetSystemMenu(WindowHandle, 0&)
        DeleteMenu ClosingButtonStyle, &HF060, &H0&
    End If
    DrawMenuBar WindowHandle
    Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
'@Ignore ProcedureNotUsed
Public Sub MinimizeForm()
    Const ProcedureName As String = "MinimizeForm"
    On Error GoTo ThrowError
    Dim TargetWindow As LongPtr
    TargetWindow = FindWindow("ThunderDFrame", Me.TargetForm.Caption)
    SendMessage TargetWindow, &H112&, &HF020&, 0
    Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub MaximizeForm()
    Const ProcedureName As String = "MaximizeForm"
    On Error GoTo ThrowError
    Dim TargetWindow As LongPtr
    TargetWindow = FindWindow("ThunderDFrame", Me.TargetForm.Caption)
    SendMessage TargetWindow, &H112&, &HF030&, 0
    Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub RestoreForm()
    Const ProcedureName As String = "RestoreForm"
    On Error GoTo ThrowError
    Dim TargetWindow As LongPtr
    TargetWindow = FindWindow("ThunderDFrame", Me.TargetForm.Caption)
    SendMessage TargetWindow, &H112&, &HF120&, 0
    Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub OnResizedForm()
    Const ProcedureName As String = "OnResizedForm"
    On Error GoTo ThrowError
    If TurnOffResizeEvent Then Exit Sub
    Select Case True
        Case IsMinimized
            Exit Sub
        Case IsMaximized
            ResizeFormControls
        Case IsNomalFromMinimization
            ResizeForm
            Exit Sub
        Case IsNomalFromMaximization
            ResizeForm
            ResizeFormControls
        Case Else
            ResizeForm
            ResizeFormControls
    End Select
    FormHeightAndWidth = CStr(Me.TargetForm.Height) & "_" & CStr(Me.TargetForm.Width)
    Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub ResizeForm()
    Const ProcedureName As String = "ResizeForm"
    On Error GoTo ThrowError
    If IsMinimized Or IsMaximized Then Exit Sub
    TurnOffResizeEvent = True
    With Me.TargetForm
        .Height = CorrectionHeight
        .Width = CorrectionWidth
    End With
    TurnOffResizeEvent = False
    Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Function CorrectionHeight() As Single
    CorrectionHeight = ReferenceValue * RATE_VERTICAL
End Function
Private Function CorrectionWidth() As Single
    CorrectionWidth = ReferenceValue * RATE_HORIZONTAL
End Function
Private Function ReferenceValue() As Long
    '@Ignore IIfSideEffect
    ReferenceValue = IIf(TemporaryCorrectionHeight < TemporaryCorrectionWidth, _
        Int(TemporaryCorrectionHeight), Int(TemporaryCorrectionWidth))
End Function
Private Function TemporaryCorrectionHeight() As Single
    TemporaryCorrectionHeight = Me.TargetForm.Height / RATE_VERTICAL
End Function
Private Function TemporaryCorrectionWidth() As Single
    TemporaryCorrectionWidth = Me.TargetForm.Width / RATE_HORIZONTAL
End Function
Private Sub ResizeFormControls()
    Const ProcedureName As String = "ResizeFormControls"
    On Error GoTo ThrowError
    Dim PositionAndSize As Dictionary
    Set PositionAndSize = StoreControlsPositionAndSize
    Dim RateVertical As Single
    Dim RateHorizontal As Single
    RateVertical = CalculateRateVertical
    RateHorizontal = CalculateRateHorizontal
    Dim KeyControl As Variant
    For Each KeyControl In PositionAndSize
        Dim Target As MSForms.Control
        Set Target = KeyControl
        Dim ItemCollection As Collection
        Set ItemCollection = PositionAndSize.Item(KeyControl)
        Dim FormSizingItem As ClassFormSizingItem
        Set FormSizingItem = ItemCollection.Item(1)
        With Target
            .Top = FormSizingItem.FormTop * RateVertical
            .Left = FormSizingItem.FormLeft * RateHorizontal
            .Height = FormSizingItem.FormHeight * RateVertical
            .Width = FormSizingItem.FormWidth * RateHorizontal
            If FormSizingItem.RateFontSize <> 0 Then
                FormExBB.LetControlFontSize Target, FormSizingItem.RateFontSize * (.Height + .Width)
            End If
        End With
    Next
    Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Function StoreControlsPositionAndSize() As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim KeyControl As MSForms.Control
    For Each KeyControl In Me.TargetForm.Controls
        With New ClassFormSizingItem
            .FormTop = KeyControl.Top
            .FormLeft = KeyControl.Left
            .FormHeight = KeyControl.Height
            .FormWidth = KeyControl.Width
            .RateFontSize = FormExBB.GetControlFontSize(KeyControl) / (KeyControl.Height + KeyControl.Width)
            Dim ItemCollection As Collection
            Set ItemCollection = New Collection
            ItemCollection.Add .FormResizingItem
        End With
        RET.Add KeyControl, ItemCollection
    Next
    Set StoreControlsPositionAndSize = RET
End Function
Private Function CalculateRateVertical() As Single
    CalculateRateVertical = Me.TargetForm.Height / BeforeHeight
End Function
Private Function CalculateRateHorizontal() As Single
    CalculateRateHorizontal = Me.TargetForm.Width / BeforeWidth
End Function
Private Function BeforeHeight() As Single
    BeforeHeight = CSng(Left$(FormHeightAndWidth, InStr(FormHeightAndWidth, "_") - 1))
End Function
Private Function BeforeWidth() As Single
    BeforeWidth = CSng(Mid$(FormHeightAndWidth, InStr(FormHeightAndWidth, "_") + 1))
End Function
Private Property Let FormHeightAndWidth(ByVal Height_Width As String)
    Me.TargetForm.Tag = Height_Width
End Property
Private Property Get FormHeightAndWidth() As String
    FormHeightAndWidth = Me.TargetForm.Tag
End Property
Private Function IsMinimized() As Boolean
    If Me.TargetForm.Height < 100 Then IsMinimized = True
End Function
Private Function IsMaximized() As Boolean
    With Application
        .WindowState = xlMaximized
        If .Height = Me.TargetForm.Height Then
            If .Width = Me.TargetForm.Width Then
                IsMaximized = True
            End If
        End If
    End With
End Function
Private Function IsNomalFromMinimization() As Boolean
    If BeforeHeight < 100 Then IsNomalFromMinimization = True
End Function
Private Function IsNomalFromMaximization() As Boolean
    With Application
        .WindowState = xlMaximized
        If .Height = BeforeHeight Then
            If .Width = BeforeWidth Then
                IsNomalFromMaximization = True
            End If
        End If
    End With
End Function


