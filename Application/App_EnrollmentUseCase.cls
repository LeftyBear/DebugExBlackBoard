VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Application.UseCase

Option Explicit

Private Type Member
    StructureResolver   As App_SchoolStructureResolver
    YearFolderResolver  As App_YearFolderResolver
    FileNameResolver    As App_FileNameResolver
    Repository          As Inf_EnrollmentRepository
    SummaryService      As Dom_EnrollmentSummaryService
End Type
Private This As Member

Public Sub Initialize(ByVal StructureResolver As App_SchoolStructureResolver, ByVal YearFolderResolver As App_YearFolderResolver, ByVal FileNameResolver As App_FileNameResolver, ByVal Repository As Inf_EnrollmentRepository, ByVal SummaryService As Dom_EnrollmentSummaryService)
    Set This.StructureResolver = StructureResolver
    Set This.YearFolderResolver = YearFolderResolver
    Set This.FileNameResolver = FileNameResolver
    Set This.Repository = Repository
    Set This.SummaryService = SummaryService
End Sub

Public Function Execute(ByVal SchoolYear As Long, ByVal TargetDate As Date, ByVal Transfer As Boolean, ByVal Remarks As String) As Dom_EnrollmentSummary
    If Not (0 < SchoolYear) Then Err.Raise App_ErrNum.NotPositiveNumber, , "SchoolYear must be positive."
    Dim Key As App_SchoolYearKey
    Set Key = New App_SchoolYearKey
    Key.Initialize SchoolYear
    Dim Structure As Dom_SchoolStructure
    Set Structure = This.StructureResolver.Resolve(Key)
        
    Dim FilePathResolver As App_EnrollmentFilePathResolver
    Set FilePathResolver = New App_EnrollmentFilePathResolver
    FilePathResolver.Initialize This.YearFolderResolver, This.FileNameResolver
    Dim FilePath As String
    FilePath = FilePathResolver.Resolve(Key)

'    This.Repository.Save FilePath, Entity
'    Dim Parser As App_EnrollmentHeaderParser
'    Set Parser = New App_EnrollmentHeaderParser
'    This.Repository.Initialize Parser
'    This.Repository.Initialize (FilePath)
    
End Function

