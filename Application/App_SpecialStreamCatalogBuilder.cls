VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SpecialStreamCatalogBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Builder")
Option Explicit

Public Function Build(ByVal Repository As Dom_IConfigRepository) As Dom_SpecialStreamCatalog
    Dim Values() As String
    Values = Repository.Load(SpecialStreamCatalog)
    Dim CSVDTO As App_CSVDTO
    Set CSVDTO = New App_CSVDTO
    CSVDTO.Initialize Values
    '@Ignore SelfAssignedDeclaration
    Dim Mapper As New App_HeaderMapper
    Dim ColumnDTOGroups As VBA.Collection
    Set ColumnDTOGroups = Mapper.Map(CSVDTO)
    Set Build = BuildSpecialStreamCatalog(ColumnDTOGroups)
End Function

Private Function BuildSpecialStreamCatalog(ByVal ColumnDTOGroups As VBA.Collection) As Dom_SpecialStreamCatalog
    Dim Result As Dom_SpecialStreamCatalog
    Set Result = New Dom_SpecialStreamCatalog
    Dim R As Long
    For R = 1 To ColumnDTOGroups.Count
        Dim ColumnDTOGroup As VBA.Collection
        Set ColumnDTOGroup = ColumnDTOGroups.Item(R)
        Dim C As Long
        For C = 1 To ColumnDTOGroup.Count
            Dim ColumnDTO As App_ColumnDTO
            Set ColumnDTO = ColumnDTOGroup.Item(C)
            Dim Index As Long
            Index = Index + 1
            Dim Name As String
            Name = ColumnDTO.RawValue
            With New Dom_SpecialStream
                Dim SpecialStream As Dom_SpecialStream
                Set SpecialStream = .Create(Index, Name)
            End With
            Result.Add SpecialStream
        Next
    Next
    Set BuildSpecialStreamCatalog = Result
End Function
