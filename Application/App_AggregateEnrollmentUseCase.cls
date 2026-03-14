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
    Parser      As Dom_EnrollmentParser
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IEnrollmentRepository, ByVal Config As Dom_SchoolConfig)
    Set This.Repository = Repository
    Set This.Config = Config
    Set This.Parser = New Dom_EnrollmentParser
End Sub

Public Function Execute(ByVal TargetDate As Date) As Dom_EnrollmentAggregate
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
    Dim Result As Dom_EnrollmentAggregate
    Set Result = New Dom_EnrollmentAggregate
    Result.Initialize Records
    Set Execute = Result
End Function

Private Function BuildRecord(ByVal ColumnDTOContainer As VBA.Collection) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim GradeCount As Long
    GradeCount = Bigger(This.Config.MainStreamGradeCount, This.Config.SpecialStreamGradeCount)
    Dim ClassCount As Long
    ClassCount = Bigger(This.Config.MainStreamClassCount, This.Config.SpecialStreamClassCount)
    Dim Values() As Long
    ReDim Values(0 To Dom_StreamType.Special, 0 To GradeCount, 0 To ClassCount, 0 To Dom_Gender.Female)
    Dim R As Long
    For R = 1 To ColumnDTOContainer.Count
        Dim Record As Dom_EnrollmentRecord
        Set Record = New Dom_EnrollmentRecord
        Dim ColumnDTOs As VBA.Collection
        Set ColumnDTOs = ColumnDTOContainer.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOs.Count
            Dim DTO As App_ColumnDTO
            Set DTO = ColumnDTOs.Item(C)
            If 0 < VBA.InStr(1, DTO.HeaderName, charUnderScore) Then
                This.Parser.Parse DTO, Values
            ElseIf DTO.HeaderName Like jpnHizuke Then
                Dim FixDate As Date
                FixDate = Dom_Function.NormalizeToDate(DTO.RawValue)
            ElseIf DTO.HeaderName Like jpnIdou Then
                Dim Transfer As Boolean
                Transfer = Dom_Function.NormalizeToBoolean(DTO.RawValue)
            ElseIf DTO.HeaderName Like jpnBikou Then
                Dim Remarks As String
                Remarks = DTO.RawValue
            End If
        Next
        Record.Initialize FixDate, Transfer, Remarks, Values
        Result.Add Record
    Next
    Set BuildRecord = Result
End Function

Private Function Bigger(ByVal Value As Long, ByVal Another As Long) As Long
    If Value < Another Then
        Bigger = Another
    Else
        Bigger = Value
    End If
End Function
