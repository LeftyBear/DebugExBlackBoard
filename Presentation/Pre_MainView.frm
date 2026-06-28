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
'@Folder "Presentation.View"
Option Explicit
Implements Pre_IViewCallback
Private Type Member
    Base            As Pre_BaseView
    Logger          As App_ILogPersistence
    UserUCFactory   As App_UserUseCaseFactory
    EditerUCFactory As App_EditerUseCaseFactory
End Type

Private This As Member

Friend Sub Inject(ByVal Base As Pre_BaseView, ByVal Logger As App_ILogPersistence, ByVal UserUCFactory As App_UserUseCaseFactory, ByVal EditerUCFactory As App_EditerUseCaseFactory)
    Set This.Base = Base
    Set This.Logger = Logger
    Set This.UserUCFactory = UserUCFactory
    Set This.EditerUCFactory = EditerUCFactory
End Sub

Public Sub OnChangeDate(ByVal SelectedDate As Date)
    ShowDailyPeriod SelectedDate
End Sub

Private Sub ShowDailyPeriod(ByVal SelectedDate As Date)
    Dim UC As App_LoadPeriodUseCase
    Set UC = This.UserUCFactory.CreateLoadPeriodUseCase
    UC.SetDate SelectedDate
    This.Base.Execute Me, UC
    With Me.ListBox1
        .ColumnCount = UBound(UC.ViewModel.Table, 2)
        .List = UC.ViewModel.Table
    End With
End Sub

Private Sub ShowSuccess(ByVal Message As String)
    If Message = vbNullString Then Exit Sub
    VBA.MsgBox Message, vbInformation, "処理完了"
End Sub

Private Sub NotifyBusinessError(ByVal Message As String)
    If Message = vbNullString Then Exit Sub
    VBA.MsgBox Message, vbExclamation, "業務エラー"
End Sub

Private Sub NotifySystemError()
    VBA.MsgBox "予期しないエラーが発生しました。", vbExclamation, "システムエラー"
End Sub

Private Sub Pre_IViewCallback_LogSystemError(ByVal Error As VBA.ErrObject)
    Dim Message As String
    Message = "ErrorNumber: " & Error.Number & _
              "  Source: " & Error.Source & _
              "  Description: " & Error.Description
    This.Logger.Log Message
End Sub

Private Sub Pre_IViewCallback_RenderResult(ByVal Result As App_UseCaseResult)
    If Result.TypeCode = SuccessCode Then
        ShowSuccess Result.Message
    ElseIf Result.TypeCode = BusinessErrorCode Then
        NotifyBusinessError Result.Message
    ElseIf Result.TypeCode = SystemErrorCode Then
        NotifySystemError
    End If
End Sub
