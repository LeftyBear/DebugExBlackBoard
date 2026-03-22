VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SubjectCatalogBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Builder")
Option Explicit

Public Function Build(ByVal Repository As Dom_IConfigRepository) As Dom_SubjectCatalog
    Dim Values() As String
    Values = Repository.Load(SubjectCatalog)
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize Values
    '@Ignore SelfAssignedDeclaration
    Dim Mapper As New App_HeaderMapper
    Dim ColumnDTOGroups As VBA.Collection
    Set ColumnDTOGroups = Mapper.Map(CSVDTO)
    Set Build = BuildSubjectCatalog(ColumnDTOGroups)
End Function

Private Function BuildSubjectCatalog(ByVal ColumnDTOGroups As VBA.Collection) As Dom_SubjectCatalog
    Dim Result As Dom_SubjectCatalog
    Set Result = New Dom_SubjectCatalog
    Dim R As Long
    For R = 1 To ColumnDTOGroups.Count
        Dim ColumnDTOGroup As VBA.Collection
        Set ColumnDTOGroup = ColumnDTOGroups.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOGroup.Count
            Dim ColumnDTO As App_ColumnDTO
            Set ColumnDTO = ColumnDTOGroup.Item(C)
            Select Case ColumnDTO.HeaderName
            Case jpnSyukeimei
                Dim Name As String
                Name = ColumnDTO.RawValue
            Case jpnTaisyoGakunen
                Dim TargetGrade As String
                TargetGrade = ColumnDTO.RawValue
            Case jpnKigou
                Dim Mark As String
                Mark = ColumnDTO.RawValue
            End Select
        Next
        Dim Index As Long
        Index = Index + 1
        With New Dom_Subject
            Dim Subject As Dom_Subject
            Set Subject = .Create(Index, Name, TargetGrade, Mark)
        End With
        Result.Add Subject
    Next
    Set BuildSubjectCatalog = Result
End Function
