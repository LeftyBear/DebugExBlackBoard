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
    StructureUC     As App_LoadStructureUseCase
    ClassHourUC     As App_LoadClassHourUseCase
    EnrollmentUC    As App_LoadEnrollmentUseCase
    ScheduleUC      As App_LoadScheduleUseCase
    ClassHourUp     As App_ClassHourUpdater
    MainView        As Pre_IMainView
    Base            As Pre_BasePresenter
End Type

Private This As Member

Public Sub Initialize(ByVal Logger As App_ILogPersistence, ByVal StructureUC As App_LoadStructureUseCase, ByVal ClassHourUC As App_LoadClassHourUseCase, ByVal EnrollmentUC As App_LoadEnrollmentUseCase, ByVal ScheduleUC As App_LoadScheduleUseCase, ByVal Base As Pre_BasePresenter)
    Set This.Logger = Logger
    Set This.StructureUC = StructureUC
    Set This.ClassHourUC = ClassHourUC
    Set This.EnrollmentUC = EnrollmentUC
    Set This.ScheduleUC = ScheduleUC
    Set This.Base = Base
End Sub

Public Sub AcceptStructure()
    Set This.Structure = This.StructureUC.Excute
End Sub

Public Sub AttachView(ByVal MainView As Pre_IMainView)
    Set This.MainView = MainView
End Sub

Public Sub LoadClassHourAsync(ByVal SchoolYear As Long)
    This.MainView.ShowLoading
    DoEvents
    This.ClassHourUp.Initialize This.ClassHourUC.Execute(SchoolYear, This.Structure)
    This.MainView.HideLoading
End Sub

Public Function LoadEnrollmentAsync(ByVal SchoolYear As Long) As Dom_IEnrollmentRepository
    This.MainView.ShowLoading
    DoEvents
    Set LoadEnrollmentAsync = This.EnrollmentUC.Execute(SchoolYear)
    This.MainView.HideLoading
End Function

Public Function LoadScheduleAsync(ByVal SchoolYear As Long) As Dom_IScheduleRepository
    This.MainView.ShowLoading
    DoEvents
    Set LoadScheduleAsync = This.ScheduleUC.Execute(SchoolYear, This.Structure)
    This.MainView.HideLoading
End Function

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
