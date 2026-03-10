VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_GenerateSchoolConfigUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Repository As Dom_IConfigRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IConfigRepository)
    Set This.Repository = Repository
End Sub

Public Function Execute() As Dom_SchoolConfig
    Dim LimitValue As Dom_LimitValue
    Set LimitValue = GenerateLimitValue
    Dim Subject As Dom_Subject
    Set Subject = GenerateSubject
    Dim Result As Dom_SchoolConfig
    Set Result = New Dom_SchoolConfig
    Result.Initialize LimitValue, Subject
    Set Execute = Result
End Function

Private Function GenerateLimitValue() As Dom_LimitValue
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(LimitValue)
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim ColumnDTOs As VBA.Collection
    Set ColumnDTOs = ColumnDTOContainer.Item(1)
    Dim C As Long
    For C = 1 To ColumnDTOs.Count
        Dim DTO As App_ColumnDTO
        Set DTO = ColumnDTOs.Item(C)
        Dim MainStreamGradeCount As Long, MainStreamClassCount As Long, SpecialGradeCount As Long, SpecialClassCount As Long, PeriodCount As Long
        ParseLimitValueColumn DTO, MainStreamGradeCount, MainStreamClassCount, SpecialGradeCount, SpecialClassCount, PeriodCount
    Next
    Dim Result As Dom_LimitValue
    Set Result = New Dom_LimitValue
    Result.Initialize MainStreamGradeCount, MainStreamClassCount, SpecialGradeCount, SpecialClassCount, PeriodCount
    Set GenerateLimitValue = Result
End Function

Private Sub ParseLimitValueColumn(ByVal DTO As App_ColumnDTO, ByRef MainStreamGradeCount As Long, ByRef MainStreamClassCount As Long, ByRef SpecialGradeCount As Long, ByRef SpecialClassCount As Long, ByRef PeriodCount As Long)
    If DTO.HeaderName Like "普通学年数" Then
        MainStreamGradeCount = CLng(DTO.RawValue)
    ElseIf DTO.HeaderName Like "普通学級数" Then
        MainStreamClassCount = CLng(DTO.RawValue)
    ElseIf DTO.HeaderName Like "特支学年数" Then
        SpecialGradeCount = CLng(DTO.RawValue)
    ElseIf DTO.HeaderName Like "特支学級数" Then
        SpecialClassCount = CLng(DTO.RawValue)
    ElseIf DTO.HeaderName Like "時間割数" Then
        PeriodCount = CLng(DTO.RawValue)
    End If
End Sub

Private Function GenerateSubject() As Dom_Subject
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(Subject)
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim Names As VBA.Collection
    Set Names = New VBA.Collection
    Dim TargetGrades As VBA.Collection
    Set TargetGrades = New VBA.Collection
    Dim R As Long
    For R = 1 To ColumnDTOContainer.Count
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = ColumnDTOContainer.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOs.Count
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            If DTO.HeaderName Like "集計名" Then
                Names.Add DTO.RawValue
            ElseIf DTO.HeaderName Like "対象学年" Then
                TargetGrades.Add DTO.RawValue
            End If
        Next
    Next
    Dim Result As Dom_Subject
    Set Result = New Dom_Subject
    Result.Initialize Names, TargetGrades
    Set GenerateSubject = Result
End Function
