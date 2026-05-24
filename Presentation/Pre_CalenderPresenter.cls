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
Implements Pre_IPresenterCallback
Private Type Member
    Structure       As Dom_SchoolStructure
    Logger          As App_ILogPersistence
    MainView        As Pre_IMainView
    Base            As Pre_BasePresenter
End Type

Private This As Member

Public Sub Initialize(ByVal Logger As App_ILogPersistence, ByVal Base As Pre_BasePresenter)
    Set This.Logger = Logger
    Set This.Base = Base
End Sub

Public Sub AttachView(ByVal MainView As Pre_IMainView)
    Set This.MainView = MainView
End Sub

Private Sub Excute()
    This.Base.Execute Me
End Sub

Private Sub Pre_IPresenterCallback_ExecuteCore()
    
End Sub

Private Sub Pre_IPresenterCallback_LogSystemError(ByVal Error As ErrObject)
    Dim Message As String
    Message = "ErrorNumber: " & Error.Number & _
              "  ModuleName: " & Error.Source & _
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
