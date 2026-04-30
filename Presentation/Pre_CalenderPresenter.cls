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
    Logger      As App_ILogger
    Structure   As App_SchoolStructureAggregator
    ClassHour   As App_ClassHourAggregator
    Enrollment  As App_EnrollmentAggregator
    Schedule    As App_ScheduleAggregator
    Factory     As Dom_ValueObjectFactory
    Builder     As App_ViewDTOBuilder
    CurrentDTO  As App_ViewDTO
    MainView    As Pre_IMainView
    Base        As Pre_BasePresenter
    CurrentYear As Long
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

Public Sub AttachView(ByVal MainView As Pre_IMainView)
    Set This.MainView = MainView
End Sub

Public Sub LoadAsync(ByVal TargetDate As Date)
    This.MainView.ShowLoading
    DoEvents
    LoadByDate TargetDate
    This.MainView.HideLoading
End Sub

Public Sub Excute()
    This.Base.Execute Me
End Sub

Private Sub LoadByDate(ByVal TargetDate As Date)
    Dim NewYear As Long
    NewYear = App_DateUtility.GetSchoolYear(TargetDate)
    If This.CurrentYear = NewYear Then Exit Sub
    Dim ClassHour As Dom_ClassHourYearAggregate
    Set ClassHour = This.ClassHour.Execute(NewYear, This.Structure)
    Dim Enrollment As Dom_EnrollmentYearAggregate
    Set Enrollment = This.Enrollment.Execute(NewYear)
    Dim Schedule As Dom_ScheduleYearAggregate
    Set Schedule = This.Schedule.Execute(NewYear, This.Structure)
    Dim Index As Long
    Index = App_DateUtility.DiffApr1st(TargetDate)
    Dim NewDTO As App_ViewDTO
    Set NewDTO = This.Builder.Build(Index, ClassHour, Enrollment, Schedule, This.Structure)
    RenderDiff This.CurrentDTO, NewDTO
    Set This.CurrentDTO = NewDTO
    This.CurrentYear = NewYear
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

Private Sub Pre_IPresenterCallback_ExecuteCore()
    
End Sub

Private Sub Pre_IPresenterCallback_LogSystemError(ByVal Error As ErrObject)
    Dim Message As String
    Message = "ErrorNumber: " & Error.Number & _
              "  ModuleName: " & Error.Source & _
              "  Description: " & Error.Description
    This.Logger.WriteLog Message
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
