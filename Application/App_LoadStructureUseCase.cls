VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadStructureUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    SchoolEventUC   As App_LoadSchoolEventUseCase
    SpecialStreamUC As App_LoadSpecialStreamUseCase
    SubjectUC       As App_LoadSubjectUseCase
    UpperValueUC    As App_LoadUpperValueUseCase
End Type

Private This As Member

Public Sub Initialize(ByVal SchoolEventUC As App_LoadSchoolEventUseCase, ByVal SpecialStreamUC As App_LoadSpecialStreamUseCase, ByVal SubjectUC As App_LoadSubjectUseCase, ByVal UpperValueUC As App_LoadUpperValueUseCase)
    Set This.SchoolEventUC = SchoolEventUC
    Set This.SpecialStreamUC = SpecialStreamUC
    Set This.SubjectUC = SubjectUC
    Set This.UpperValueUC = UpperValueUC
End Sub

Public Function Excute() As Dom_SchoolStructure
    Dim SchoolEventRepo As Dom_ISchoolEventRepository
    Set SchoolEventRepo = This.SchoolEventUC.Execute
    Dim SpecialStreamRepo As Dom_ISpecialStreamRepository
    Set SpecialStreamRepo = This.SpecialStreamUC.Execute
    Dim SubjectPero As Dom_ISubjectRepository
    Set SubjectPero = This.SubjectUC.Execute
    Dim UpperValueRepo As Dom_IUpperValueRepository
    Set UpperValueRepo = This.UpperValueUC.Execute
    Dim Result As Dom_SchoolStructure
    Set Result = New Dom_SchoolStructure
    Result.Initialize SchoolEventRepo, SpecialStreamRepo, SubjectPero, UpperValueRepo
    Set Excute = Result
End Function
