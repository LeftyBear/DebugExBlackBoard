VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_AggregateClassHourUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Private Type Member
    Repository  As Dom_IClassHourRepository
    Config      As Dom_SchoolConfig
    ClassHour   As Dom_ClassHourParser
    TimeTable   As Dom_TimeTableParser
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IClassHourRepository, ByVal Config As Dom_SchoolConfig)
    Set This.Repository = Repository
    Set This.Config = Config
    Set This.ClassHour = New Dom_ClassHourParser
    Set This.TimeTable = New Dom_TimeTableParser
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_ClassHourAggregate
    Dim SchoolYear As Long
    SchoolYear = Dom_Function.GetSchoolYear(TargetDate)
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize This.Repository.Load(SchoolYear)
    If CSVDTO.RowCount = 0 Then Exit Function
    Dim AddColumnDTO As App_AddColumnDTOatRowUseCase
    Set AddColumnDTO = New App_AddColumnDTOatRowUseCase
    Dim ColumnDTOContainer As VBA.Collection
    Set ColumnDTOContainer = AddColumnDTO.Execute(CSVDTO)
    Dim Records As VBA.Collection
    Set Records = BuildRecord(ColumnDTOContainer)
    Dim Result As Dom_ClassHourAggregate
    Set Result = New Dom_ClassHourAggregate
    Result.Initialize Records
    Set Execute = Result
End Function

Private Function BuildRecord(ByVal ColumnDTOContainer As VBA.Collection) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Values() As Long
    ReDim Values(0 To Dom_ClassHourType.Execution, 0 To This.Config.SubjectCount, 0 To This.Config.MainStreamGradeCount)
    Dim TimeTables() As String
    ReDim TimeTables(0 To Dom_ClassHourType.Execution, 0 To This.Config.PeriodCount, 0 To This.Config.MainStreamGradeCount)
    Dim R As Long
    For R = 1 To ColumnDTOContainer.Count
        Dim Record As Dom_ClassHourRecord
        Set Record = New Dom_ClassHourRecord
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = ColumnDTOContainer.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOs.Count
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            If 0 < VBA.InStr(1, DTO.HeaderName, charUnderScore) Then
                If 0 < VBA.InStr(1, DTO.HeaderName, jpnKouji) Then
                    This.TimeTable.Parse This.Config, DTO, TimeTables
                Else
                    This.ClassHour.Parse This.Config, DTO, Values
                End If
            ElseIf DTO.HeaderName Like jpnHizuke Then
                Dim FixDate As Date
                FixDate = Dom_Function.NormalizeToDate(DTO.RawValue)
            End If
        Next
        Record.Initialize FixDate, Values, TimeTables
        Result.Add Record
    Next
    Set BuildRecord = Result
End Function
