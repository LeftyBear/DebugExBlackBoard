VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_Presenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.ViewModel")
Option Explicit
Implements App_IPresenterCallback
Private Type Member
    CurrentYear As Dom_SchoolYear
    Logger      As App_ILogger
    Structure   As App_SchoolStructureAggregator
    ClassHour   As App_ClassHourAggregator
    Enrollment  As App_EnrollmentAggregator
    Schedule    As App_ScheduleAggregator
    Factory     As Dom_ValueObjectFactory
    Builder     As App_ViewDTOBuilder
    CurrentDTO  As App_ViewDTO
    MainView    As App_IMainView
    Base        As Pre_BasePresenter
End Type

Private This As Member

Public Sub Initialize(ByVal Logger As App_ILogger, ByVal Structure As App_SchoolStructureAggregator, ByVal ClassHour As App_ClassHourAggregator, ByVal Enrollment As App_EnrollmentAggregator, ByVal Schedule As App_ScheduleAggregator, ByVal Factory As Dom_ValueObjectFactory, ByVal Builder As App_ViewDTOBuilder, ByVal Base As Pre_BasePresenter)
    Set This.Logger = Logger
    Set This.Structure = Structure
    Set This.ClassHour = ClassHour
    Set This.Enrollment = Enrollment
    Set This.Schedule = Schedule
    Set This.Factory = Factory
    Set This.Builder = Builder
    Set This.Base = Base
End Sub

Public Sub AttachView(ByVal MainView As App_IMainView)
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

Private Sub LoadByDate(ByVal SelectedDate As Date)
    Dim SchoolYear As Long
    SchoolYear = App_DateUtility.GetSchoolYear(SelectedDate)
    Dim NewYear As Dom_SchoolYear
    Set NewYear = This.Factory.CreateSchoolYear(SchoolYear)
    If This.CurrentYear.Equals(NewYear) Then Exit Sub
    Dim ClassHour As Dom_ClassHourYearAggregate
    Set ClassHour = This.ClassHour.Execute(NewYear, This.Structure)
    Dim Enrollment As Dom_EnrollmentYearAggregate
    Set Enrollment = This.Enrollment.Execute(NewYear)
    Dim Schedule As Dom_ScheduleYearAggregate
    Set Schedule = This.Schedule.Execute(NewYear, This.Structure)
    Dim TargetDate As Dom_DateID
    Set TargetDate = This.Factory.CreateDateID(SelectedDate)
    Dim NewDTO As App_ViewDTO
    Set NewDTO = This.Builder.Build(TargetDate, ClassHour, Enrollment, Schedule, This.Structure)
    RenderDiff This.CurrentDTO, NewDTO
    Set This.CurrentDTO = NewDTO
    Set This.CurrentYear = NewYear
End Sub

Private Sub RenderDiff(ByVal OldDTO As App_ViewDTO, ByVal NewDTO As App_ViewDTO)
    If OldDTO Is Nothing Then
        This.MainView.Render NewDTO
        Exit Sub
    End If
    If Not IsSameArray(OldDTO.EnrollmentTable, NewDTO.EnrollmentTable) Then
        This.MainView.RenderEnrollment NewDTO.EnrollmentTable
    End If
    If Not IsSameArray(OldDTO.ClassHourPlanTable, NewDTO.ClassHourPlanTable) Then
        This.MainView.RenderClassHourPlan NewDTO.ClassHourPlanTable
    End If
    If Not IsSameArray(OldDTO.ClassHourExecutionTable, NewDTO.ClassHourExecutionTable) Then
        This.MainView.RenderClassHourExecution NewDTO.ClassHourExecutionTable
    End If
    If Not IsSameArray(OldDTO.TimeTablePlanTable, NewDTO.TimeTablePlanTable) Then
        This.MainView.RenderTimeTablePlan NewDTO.TimeTablePlanTable
    End If
    If Not IsSameArray(OldDTO.TimeTableExecutionTable, NewDTO.TimeTableExecutionTable) Then
        This.MainView.RenderTimeTableExecution NewDTO.TimeTableExecutionTable
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

Private Sub App_IPresenterCallback_ExecuteCore()
    
End Sub

Private Sub App_IPresenterCallback_LogSystemError(ByVal Error As ErrObject)
    Dim Message As String
    Message = "ErrorNumber: " & Error.Number & _
              "  ModuleName: " & Error.Source & _
              "  Description: " & Error.Description
    This.Logger.WriteLog Message
End Sub

Private Sub App_IPresenterCallback_RenderResult(ByVal Result As Pre_PresenterResult)
    Select Case Result.ResultType
    Case Pre_ViewResultTypeEnum.Success
        This.MainView.ShowSuccess Result.Message
    Case Pre_ViewResultTypeEnum.BusinessError
        This.MainView.NotifyBusinessError Result.Message
    Case Pre_ViewResultTypeEnum.SystemError
        This.MainView.NotifySystemError
    End Select
End Sub


