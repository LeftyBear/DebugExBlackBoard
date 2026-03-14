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
    Dim MainStream As Dom_MainStream
    Set MainStream = GenerateMainStream
    Dim SpecialStream As Dom_SpecialStream
    Set SpecialStream = GenerateSpecialClass
    Dim Subject As Dom_Subject
    Set Subject = GenerateSubject
    Dim TimeTable As Dom_TimeTable
    Set TimeTable = GenerateTimeTable
    Dim Result As Dom_SchoolConfig
    Set Result = New Dom_SchoolConfig
    Result.Initialize MainStream, SpecialStream, Subject, TimeTable
    Set Execute = Result
End Function

Private Function GenerateMainStream() As Dom_MainStream
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(MainStream)
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
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            Select Case DTO.HeaderName
            Case jpnFutsuGakunenSu
                Dim GradeCount As Long
                GradeCount = CLng(DTO.RawValue)
            Case jpnFutsuGakkyuSu
                Dim ClassCount As Long
                ClassCount = CLng(DTO.RawValue)
            End Select
        Next
    Next
    Dim Result As Dom_MainStream
    Set Result = New Dom_MainStream
    Result.Initialize GradeCount, ClassCount
    Set GenerateMainStream = Result
End Function

Private Function GenerateSpecialClass() As Dom_SpecialStream
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(SpecialStream)
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim R As Long
    For R = 1 To ColumnDTOContainer.Count
        Dim Names() As String
        ReDim Names(ColumnDTOContainer.Count)
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = ColumnDTOContainer.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOs.Count
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            Names(R) = DTO.RawValue
        Next
    Next
    Dim Result As Dom_SpecialStream
    Set Result = New Dom_SpecialStream
    Result.Initialize Names
    Set GenerateSpecialClass = Result
End Function

Private Function GenerateSubject() As Dom_Subject
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(Subject)
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim Names() As String
    ReDim Names(ColumnDTOContainer.Count)
    Dim TargetGrades() As String
    ReDim TargetGrades(ColumnDTOContainer.Count)
    Dim Marks() As String
    ReDim Marks(ColumnDTOContainer.Count)
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
                Names(R) = DTO.RawValue
            Case jpnTaisyoGakunen
                TargetGrades(R) = DTO.RawValue
            Case jpnKigou
                Marks(R) = DTO.RawValue
            End Select
        Next
    Next
    Dim Result As Dom_Subject
    Set Result = New Dom_Subject
    Result.Initialize Names, TargetGrades, Marks
    Set GenerateSubject = Result
End Function

Private Function GenerateTimeTable() As Dom_TimeTable
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(TimeTable)
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
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            Dim PeriodCount As Long
            PeriodCount = CLng(DTO.RawValue)
        Next
    Next
    Dim Result As Dom_TimeTable
    Set Result = New Dom_TimeTable
    Result.Initialize PeriodCount
    Set GenerateTimeTable = Result
End Function
