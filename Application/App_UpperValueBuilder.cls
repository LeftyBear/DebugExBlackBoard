VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_UpperValueBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Builder")
Option Explicit

Public Function Build(ByVal Repository As Dom_IConfigRepository) As Dom_UpperValue
    Dim Values() As String
    Values = Repository.Load(UpperValue)
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize Values
    '@Ignore SelfAssignedDeclaration
    Dim Mapper As New App_HeaderMapper
    Dim ColumnDTOGroups As VBA.Collection
    Set ColumnDTOGroups = Mapper.Map(CSVDTO)
    Set Build = BuildUpperValue(ColumnDTOGroups)
End Function

Private Function BuildUpperValue(ByVal ColumnDTOGroups As VBA.Collection) As Dom_UpperValue
    Dim R As Long
    For R = 1 To ColumnDTOGroups.Count
        Dim ColumnDTOGroup As VBA.Collection
        Set ColumnDTOGroup = ColumnDTOGroups.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOGroup.Count
            Dim ColumnDTO As App_ColumnDTO
            Set ColumnDTO = ColumnDTOGroup.Item(C)
            Select Case ColumnDTO.HeaderName
            Case jpnFutsuGakunenSu
                Dim MainStreamGradeCount As Long
                MainStreamGradeCount = Dom_Util_Long.NormalizeToLong(ColumnDTO.RawValue)
            Case jpnFutsuGakkyuSu
                Dim MainStreamClassCount As Long
                MainStreamClassCount = Dom_Util_Long.NormalizeToLong(ColumnDTO.RawValue)
            Case jpnJikanwariSu
                Dim PeriodCount As Long
                PeriodCount = Dom_Util_Long.NormalizeToLong(ColumnDTO.RawValue)
            End Select
        Next
        With New Dom_UpperValue
            Dim Result As Dom_UpperValue
            Set Result = .Create(MainStreamGradeCount, MainStreamClassCount, PeriodCount)
        End With
    Next
    Set BuildUpperValue = Result
End Function
