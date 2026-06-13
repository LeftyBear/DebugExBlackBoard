VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_CalenderPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.ViewModel")
Option Explicit
Implements App_IUseCaseCallback
Private Type Member
    Logger      As App_ILogPersistence
    MainView    As Pre_IMainView
    Base        As App_BaseUseCase
    TargetDate  As Date
End Type

Private This As Member

Friend Sub Initialize(ByVal Logger As App_ILogPersistence, ByVal Base As App_BaseUseCase)
    Set This.Logger = Logger
    Set This.Base = Base
End Sub

Friend Sub Inject(ByVal MainView As Pre_IMainView)
    Set This.MainView = MainView
End Sub

Public Sub Present(ByVal TargetDate As Date)
    This.TargetDate = TargetDate
    This.Base.Execute Me
End Sub

Private Sub App_IUseCaseCallback_ExecuteCore()
    'RepaintCalenderUseCase
End Sub

Private Sub App_IUseCaseCallback_LogSystemError(ByVal Error As ErrObject)
    Dim Message As String
    Message = "ErrorNumber: " & Error.Number & _
              "  ModuleName: " & Error.Source & _
              "  Description: " & Error.Description
    This.Logger.Log Message
End Sub

Private Sub App_IUseCaseCallback_RenderResult(ByVal Result As App_UseCaseResult)
    If Result.TypeCode = Success Then
        This.MainView.ShowSuccess Result.Message
    ElseIf Result.TypeCode = BusinessError Then
        This.MainView.NotifyBusinessError Result.Message
    ElseIf Result.TypeCode = SystemError Then
        This.MainView.NotifySystemError
    End If
End Sub
