VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateEnrollmentUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Repository  As Dom_IEnrollmentRepository
    Config      As Dom_SchoolConfig
    Parser      As Dom_EnrollmentHeaderParser
    Interpreter As Dom_EnrollmentInterpreter
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IEnrollmentRepository, ByVal Config As Dom_SchoolConfig, ByVal Parser As Dom_EnrollmentHeaderParser, ByVal Interpreter As Dom_EnrollmentInterpreter)
    Set This.Repository = Repository
    Set This.Config = Config
    Set This.Parser = Parser
    Set This.Interpreter = Interpreter
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_EnrollmentYearAggregate
    Dim SchoolYear As Long
    SchoolYear = Dom_Util_Date.GetSchoolYear(TargetDate)
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(SchoolYear)
    If CSVDTO.RowCount = 0 Then Exit Function
    With New App_AddColumnDTOatRowUseCase
        Dim ColumnDTOContainer As VBA.Collection
        Set ColumnDTOContainer = .Execute(CSVDTO)
    End With
    Dim Factory As Dom_EnrollmentAggregateFactory
    Set Factory = New Dom_EnrollmentAggregateFactory
    Factory.Initialize This.Parser, This.Interpreter
    Set Execute = Factory.Create(This.Config, ColumnDTOContainer)
End Function
