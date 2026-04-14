VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_Presenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Presenter"
Option Explicit
Implements App_IPresenterCallback
Private Type Member
    Structure   As Dom_SchoolStructure
    VOFactory   As Dom_ValueObjectFactory
    Logger      As App_ILogger
    Schedule    As App_AggregateSchedule
    Enrollment  As App_AggregateEnrollment
    ClassHour   As App_AggregateClassHour
    Builder     As Pre_ViewModelBuilder
    CurrentVM   As Pre_ViewModel
    Base        As Pre_BasePresenter
    MainView    As Pre_IMainView
End Type

Private This As Member

Public Sub Initialize(ByVal Structure As Dom_SchoolStructure, ByVal Logger As App_ILogger, ByVal Schedule As App_AggregateSchedule, ByVal Enrollment As App_AggregateEnrollment, ByVal ClassHour As App_AggregateClassHour)
    Set This.Structure = Structure
    Set This.VOFactory = New Dom_ValueObjectFactory
    Set This.Logger = Logger
    Set This.Schedule = Schedule
    Set This.Enrollment = Enrollment
    Set This.ClassHour = ClassHour
    Set This.Builder = New Pre_ViewModelBuilder
    Set This.Base = New Pre_BasePresenter
End Sub

Public Sub AttachView(ByVal MainView As Pre_IMainView)
    Set This.MainView = MainView
End Sub

Public Sub LoadAsync(ByVal SelectedDate As Date)
    This.MainView.ShowLoading
    DoEvents
    LoadByDate SelectedDate
    This.MainView.HideLoading
End Sub

Public Sub Excute()
    This.Base.Execute Me
End Sub

Private Sub App_IPresenterCallback_ExecuteCore()
    
End Sub

Private Sub App_IPresenterCallback_LogSystemError(ByVal Error As ErrObject)
    Dim Message As String
    Message = "ErrorNumber: " & Error.Number & _
              "  ModuleName: " & Error.Source & _
              "  Description: " & Error.Description
    This.Logger.WriteLog Message
End Sub

Private Sub App_IPresenterCallback_RenderResult(ByVal Result As App_PresenterResult)
    Select Case Result.ResultType
    Case App_ViewResultType.Success
        This.MainView.ShowSuccess Result.Message
    Case App_ViewResultType.BusinessError
        This.MainView.NotifyBusinessError Result.Message
    Case App_ViewResultType.SystemError
        This.MainView.NotifySystemError
    End Select
End Sub

Private Sub LoadByDate(ByVal SelectedDate As Date)
    Dim ScheduleAgg As Dom_ScheduleYearAggregate
    Set ScheduleAgg = This.Schedule.Execute(SelectedDate, This.Structure)
    Dim EnrollmentAgg As Dom_EnrollmentYearAggregate
    Set EnrollmentAgg = This.Enrollment.Execute(SelectedDate)
    Dim ClassHourAgg As Dom_ClassHourYearAggregate
    Set ClassHourAgg = This.ClassHour.Execute(SelectedDate, This.Structure)
    Dim TargetDate As Dom_DateID
    Set TargetDate = This.VOFactory.CreateDateID(SelectedDate)
    Dim Schedule As Dom_ScheduleAggregate
    Set Schedule = ScheduleAgg.GetByDate(TargetDate)
    Dim NewVM As Pre_ViewModel
    Set NewVM = This.Builder.Build(Schedule, EnrollmentAgg, ClassHourAgg, This.Structure, TargetDate)
    RenderDiff This.CurrentVM, NewVM
    Set This.CurrentVM = NewVM
End Sub

Private Sub RenderDiff(ByVal OldVM As Pre_ViewModel, ByVal NewVM As Pre_ViewModel)
    If OldVM Is Nothing Then
        This.MainView.Render NewVM
        Exit Sub
    End If
    If Not IsSameArray(OldVM.EnrollmentTable, NewVM.EnrollmentTable) Then
        This.MainView.RenderEnrollment NewVM.EnrollmentTable
    End If
    If Not IsSameArray(OldVM.ClassHourPlanTable, NewVM.ClassHourPlanTable) Then
        This.MainView.RenderClassHourPlan NewVM.ClassHourPlanTable
    End If
    If Not IsSameArray(OldVM.ClassHourExecutionTable, NewVM.ClassHourExecutionTable) Then
        This.MainView.RenderClassHourExecution NewVM.ClassHourExecutionTable
    End If
    If Not IsSameArray(OldVM.TimeTablePlanTable, NewVM.TimeTablePlanTable) Then
        This.MainView.RenderTimeTablePlan NewVM.TimeTablePlanTable
    End If
    If Not IsSameArray(OldVM.TimeTableExecutionTable, NewVM.TimeTableExecutionTable) Then
        This.MainView.RenderTimeTableExecution NewVM.TimeTableExecutionTable
    End If
End Sub

Private Function IsSameArray(ByRef OneArray As Variant, ByRef AnotherArray As Variant) As Boolean
    On Error GoTo ErrHandle
    If UBound(OneArray) <> UBound(AnotherArray) Then Exit Function
    Dim i As Long
    For i = LBound(OneArray) To UBound(OneArray)
        If OneArray(i) <> AnotherArray(i) Then Exit Function
    Next
    IsSameArray = True
    Exit Function
ErrHandle:
    IsSameArray = False
End Function
