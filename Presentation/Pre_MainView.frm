VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pre_MainView 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "Pre_MainView.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Pre_MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Application.View"
Option Explicit
Implements Pre_IMainView
Private Type Member
    ClassHourController As Pre_ClassHourController
End Type

Private This As Member

Public Sub Initialize(ByVal ClassHourController As Pre_ClassHourController)
    Set This.ClassHourController = ClassHourController
End Sub

Private Sub DateChanged(ByVal SelectedDate As Date)
    This.ClassHourController.ChangeDate SelectedDate
End Sub

Private Sub SetGridValue(ByVal Kind As String, ByVal Value As Variant, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long)
    Dim Cell As Object
    Set Cell = ResolveGridControl(Kind, Grade, ClassNo)
    If Cell Is Nothing Then Exit Sub
    Cell.Text = CStr(Value)
End Sub

Private Function GetGridLongValue(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As Long
    Dim TextValue As String
    TextValue = GetTextFromGrid(Kind, Grade, ClassNo)
    If TextValue = vbNullString Then
        GetGridLongValue = 0
        Exit Function
    End If
    GetGridLongValue = CLng(TextValue)
End Function

Private Function GetGridStringValue(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As String
    GetGridStringValue = GetTextFromGrid(Kind, Grade, ClassNo)
End Function

Private Function GetTextFromGrid(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As String
    Dim Control As Object
    Set Control = ResolveGridControl(Kind, Grade, ClassNo)
    If Control Is Nothing Then
        GetTextFromGrid = vbNullString
        Exit Function
    End If
    GetTextFromGrid = CStr(Control.Text)
End Function

Private Function ResolveGridControl(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As Object
    Dim ControlName As String
    ControlName = BuildGridControlName(Kind, Grade, ClassNo)
    On Error Resume Next
    Set ResolveGridControl = Me.Controls(ControlName)
    On Error GoTo 0
End Function

Private Function BuildGridControlName(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As String
    Dim Cells() As Variant
    Cells = Array(Kind, CStr(Grade), CStr(ClassNo))
    BuildGridControlName = VBA.Join(Cells, charUnderScore)
End Function

Private Sub Pre_IMainView_HideLoading()
    Application.StatusBar = ""
End Sub

Private Sub Pre_IMainView_NotifyBusinessError(ByVal Message As String)
    MsgBox Message, vbCritical, "業務エラー"
End Sub

Private Sub Pre_IMainView_NotifySystemError()
    MsgBox "予期しないエラーが発生したのでログに書き出しました。", vbCritical, "システムエラー"
End Sub

Private Sub Pre_IMainView_Render(ByVal VM As Pre_ViewModel)
    Pre_IMainView_RenderEnrollment VM.EnrollmentTable
    Pre_IMainView_RenderClassHourPlan VM.ClassHourPlanTable
    Pre_IMainView_RenderClassHourExecution VM.ClassHourExecutionTable
    Pre_IMainView_RenderTimeTablePlan VM.TimeTablePlanTable
    Pre_IMainView_RenderTimeTableExecution VM.TimeTableExecutionTable
End Sub

Private Sub Pre_IMainView_RenderClassHourExecution(Table() As Variant)

End Sub

Private Sub Pre_IMainView_RenderClassHourPlan(Table() As Variant)

End Sub

Private Sub Pre_IMainView_RenderEnrollment(Table() As Variant)

End Sub

Private Sub Pre_IMainView_RenderTimeTableExecution(Table() As Variant)

End Sub

Private Sub Pre_IMainView_RenderTimeTablePlan(Table() As Variant)

End Sub

Private Sub Pre_IMainView_ShowLoading()
    Application.StatusBar = "Loading..."
    DoEvents
End Sub

Private Sub Pre_IMainView_ShowSuccess(ByVal Message As String)
    If Message = vbNullString Then Exit Sub
    MsgBox Message, vbInformation, "処理完了"
End Sub
