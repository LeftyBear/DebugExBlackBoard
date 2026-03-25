VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateEnrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Builder     As App_EntityFilePathBuilder
    Repository  As Dom_IEnrollmentRepository
    Structure   As Dom_SchoolStructure
    Parser      As Dom_EnrollmentHeaderParser
    Interpreter As Dom_EnrollmentInterpreter
End Type

Private This As Member

Public Sub Initialize(ByVal Builder As App_EntityFilePathBuilder, ByVal Repository As Dom_IEnrollmentRepository, ByVal Structure As Dom_SchoolStructure, ByVal Parser As Dom_EnrollmentHeaderParser, ByVal Interpreter As Dom_EnrollmentInterpreter)
    Set This.Builder = Builder
    Set This.Repository = Repository
    Set This.Structure = Structure
    Set This.Parser = Parser
    Set This.Interpreter = Interpreter
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_EnrollmentYearAggregate
    With New Dom_SchoolYear
        Dim SchoolYear As Dom_SchoolYear
        Set SchoolYear = .Create(Dom_Util_Date.GetSchoolYear(TargetDate))
    End With
    Dim FilePath As String
    FilePath = This.Builder.Builed(Enrollment, SchoolYear)
    If VBA.Len(VBA.Dir(FilePath)) = 0 Then Err.Raise AppErrInvalidFilePath, "App_AggregateEnrollment", "File path is invalid."
'    Dim CSVDTO As App_CSVDTO
'    Set CSVDTO = New App_CSVDTO
'    CSVDTO.Initialize This.Repository.Load(SchoolYear)
'    If CSVDTO.RowCount = 0 Then Err.Raise AppErrEmptyArray, "App_AggregateEnrollment", "CSV data is empty."
'    With New App_HeaderMapper
'        Dim EntityGroups As VBA.Collection
'        Set EntityGroups = .Map(CSVDTO)
'    End With
'    Dim Factory As Dom_EnrollmentAggregateFactory
'    Set Factory = New Dom_EnrollmentAggregateFactory
'    Factory.Initialize This.Parser, This.Interpreter
'    Set Execute = Factory.Create(This.Structure, EntityGroups)
End Function
