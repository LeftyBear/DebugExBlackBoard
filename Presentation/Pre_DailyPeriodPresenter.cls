VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_DailyPeriodPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presenter.ViewModel")
Option Explicit
Implements Pre_IPresenterCallback
Private Type Member
    UseCase     As App_BuildPeriodTotalUseCase
    Logger      As App_ILogPersistence
    Base        As Pre_BasePresenter
    MainView    As Pre_IMainView
    TargetDate  As Date
End Type

Private This As Member

Friend Sub Inject(ByVal UseCase As App_BuildPeriodTotalUseCase, ByVal Logger As App_ILogPersistence)
    Set This.UseCase = UseCase
    Set This.Logger = Logger
    Set This.Base = New Pre_BasePresenter
End Sub

Friend Sub AttachView(ByVal MainView As Pre_IMainView)
    Set This.MainView = MainView
End Sub

Public Sub ViewDailyPeriodTotal(ByVal TargetDate As Date)
    This.TargetDate = TargetDate
    This.Base.Execute Me
End Sub

Private Sub Pre_IPresenterCallback_ExecuteCore()
    Dim ViewTable() As Variant
    ViewTable = This.UseCase.Execute(This.TargetDate)
    
End Sub

Private Sub Pre_IPresenterCallback_LogSystemError(ByVal Error As VBA.ErrObject)
    Dim Message As String
    Message = "ErrorNumber: " & Error.Number & _
              "  Source: " & Error.Source & _
              "  Description: " & Error.Description
    This.Logger.Log Message
End Sub

Private Sub Pre_IPresenterCallback_RenderResult(ByVal Result As Pre_PresenterResult)
    If Result.TypeCode = Pre_ViewResultTypePolicy.Success Then
        This.MainView.ShowSuccess Result.Message
    ElseIf Result.TypeCode = Pre_ViewResultTypePolicy.BusinessError Then
        This.MainView.NotifyBusinessError Result.Message
    ElseIf Result.TypeCode = Pre_ViewResultTypePolicy.SystemError Then
        This.MainView.NotifySystemError
    End If
End Sub
