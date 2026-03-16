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
    Dim UpperValue As Dom_UpperValue
    Set UpperValue = GenerateUpperValue
    Dim SpecialStreamCatalog As Dom_SpecialStreamCatalog
    Set SpecialStreamCatalog = GenerateSpecialStreamCatalog
    Dim SubjectCatalog As Dom_SubjectCatalog
    Set SubjectCatalog = GenerateSubjectCatalog
    Dim Result As Dom_SchoolConfig
    Set Result = New Dom_SchoolConfig
    Result.Initialize UpperValue, SpecialStreamCatalog, SubjectCatalog
    Set Execute = Result
End Function

Private Function GenerateUpperValue() As Dom_UpperValue
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(UpperValue)
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim R As Long
    For R = 1 To ColumnDTOContainer.Count
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = ColumnDTOContainer.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOs.Count
            Dim MainStreamGradeCount As Long, MainStreamClassCount As Long, PeriodCount As Long
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            Select Case DTO.HeaderName
            Case jpnFutsuGakunenSu
                MainStreamGradeCount = CLng(DTO.RawValue)
            Case jpnFutsuGakkyuSu
                MainStreamClassCount = CLng(DTO.RawValue)
            Case jpnJikanwariSu
                PeriodCount = CLng(DTO.RawValue)
            End Select
        Next
        Dim Result As Dom_UpperValue
        Set Result = New Dom_UpperValue
        Result.Initialize MainStreamGradeCount, MainStreamClassCount, PeriodCount
    Next
    Set GenerateUpperValue = Result
End Function

Private Function GenerateSpecialStreamCatalog() As Dom_SpecialStreamCatalog
    Dim Result As Dom_SpecialStreamCatalog
    Set Result = New Dom_SpecialStreamCatalog
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(SpecialStreamCatalog)
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim R As Long
    For R = 1 To ColumnDTOContainer.Count
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = ColumnDTOContainer.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOs.Count
            Dim SpecialStream As Dom_SpecialStream
            Set SpecialStream = New Dom_SpecialStream
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            Dim Index As Long
            Index = Index + 1
            Dim Name As String
            Name = DTO.RawValue
            SpecialStream.Initialize Index, Name
            Result.Add SpecialStream
        Next
    Next
    Set GenerateSpecialStreamCatalog = Result
End Function

Private Function GenerateSubjectCatalog() As Dom_SubjectCatalog
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(SubjectCatalog)
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim Result As Dom_SubjectCatalog
    Set Result = New Dom_SubjectCatalog
    Dim R As Long
    For R = 1 To ColumnDTOContainer.Count
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = ColumnDTOContainer.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOs.Count
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            Select Case DTO.HeaderName
            Case jpnSyukeimei
                Dim Name As String
                Name = DTO.RawValue
            Case jpnTaisyoGakunen
                Dim TargetGrade As String
                TargetGrade = DTO.RawValue
            Case jpnKigou
                Dim Mark As String
                Mark = DTO.RawValue
            End Select
        Next
        Dim Index As Long
        Index = Index + 1
        Dim Subject As Dom_Subject
        Set Subject = New Dom_Subject
        Subject.Initialize Index, Name, TargetGrade, Mark
        Result.Add Subject
    Next
    Set GenerateSubjectCatalog = Result
End Function
