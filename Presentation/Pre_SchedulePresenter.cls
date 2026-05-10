VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_SchedulePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.ViewModel")
Option Explicit
Implements Pre_IPresenterCallback
Private Type Member
    Logger      As App_ILogPersistence
    Factory     As App_ViewDTOFactory
    MainView    As Pre_IMainView
    Base        As Pre_BasePresenter
    CurrentDTO  As App_ViewDTO
End Type

Private This As Member

Public Sub Initialize(ByVal Logger As App_ILogPersistence, ByVal Factory As App_ViewDTOFactory, ByVal Base As Pre_BasePresenter)
    Set This.Logger = Logger
    Set This.Factory = Factory
    Set This.Base = Base
End Sub

Public Sub AttachView(ByVal MainView As Pre_IMainView)
    Set This.MainView = MainView
End Sub

Private Sub Excute()
    This.Base.Execute Me
End Sub

Public Sub Show(ByVal Index As Long, ByVal ClassHour As VBA.Collection, ByVal Enrollment As VBA.Collection, ByVal Schedule As VBA.Collection, ByVal Structure As Dom_SchoolStructure)
    Dim NewDTO As App_ViewDTO
    Set NewDTO = This.Factory.Create(Index, ClassHour, Enrollment, Schedule, Structure)
    RenderDiff This.CurrentDTO, NewDTO
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

Private Function IsSameArray(ByRef OneArray() As Variant, ByRef AnotherArray() As Variant) As Boolean
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
