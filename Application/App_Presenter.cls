VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_Presenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.Presenter"
Option Explicit
Implements App_IPresenterCallback
Private Type Member
    Logger      As Inf_ILogger
    Structure   As Dom_SchoolStructure
    Enrollment  As App_AggregateEnrollment
    ClassHour   As App_AggregateClassHour
    MainView    As App_IMainView
    Base        As App_BasePresenter
End Type

Private This As Member

Public Sub Initialize(ByVal Logger As Inf_ILogger, ByVal Structure As Dom_SchoolStructure, ByVal Enrollment As App_AggregateEnrollment, ByVal ClassHour As App_AggregateClassHour)
    Set This.Logger = Logger
    Set This.Structure = Structure
    Set This.Enrollment = Enrollment
    Set This.ClassHour = ClassHour
    Set This.Base = New App_BasePresenter
End Sub

Public Sub AttachView(ByVal MainView As App_IMainView)
    Set This.MainView = MainView
End Sub

Public Sub Bootstrap()
    Dim YearlyEnrollment As Dom_EnrollmentAggregate
    Set YearlyEnrollment = This.Enrollment.Execute(Date)
    Dim YearlyClassHour As Dom_ClassHourAggregate
    Set YearlyClassHour = This.ClassHour.Execute(Date, This.Structure)
    'View----------------------------------------------------------------------------------------
    Dim EnrollmentFormatter As App_ViewEnrollmentFormatter
    Set EnrollmentFormatter = New App_ViewEnrollmentFormatter
    Dim EnrollmentTable() As Variant
    EnrollmentTable = EnrollmentFormatter.Format(YearlyEnrollment.GetByDate(Date), This.Structure)
    Dim ClassHourFormatter As App_ViewClassHourFormatter
    Set ClassHourFormatter = New App_ViewClassHourFormatter
    Dim ClassHourPlanTable() As Variant
    ClassHourPlanTable = ClassHourFormatter.Format(Plan, YearlyClassHour.GetByDate(Date), This.Structure)
    Dim ClassHourExecutionTable() As Variant
    ClassHourExecutionTable = ClassHourFormatter.Format(Execution, YearlyClassHour.GetByDate(Date), This.Structure)
    Dim TimeTableFormatter As App_ViewTimeTableFormatter
    Set TimeTableFormatter = New App_ViewTimeTableFormatter
    Dim TimeTablePlanTable() As Variant
    TimeTablePlanTable = TimeTableFormatter.Format(Plan, YearlyClassHour.GetByDate(Date), This.Structure)
    Dim TimeTableExecutionTable() As Variant
    TimeTableExecutionTable = TimeTableFormatter.Format(Execution, YearlyClassHour.GetByDate(Date), This.Structure)
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
