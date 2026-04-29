VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ConfigFileNameResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Service")
Option Explicit

Public Function Resolve(ByVal StructureName As Dom_StructureTypeEnum) As String
    If StructureName = Dom_StructureTypeEnum.SchoolEvent Then
        Resolve = "SchoolEventCatalog.csv"
    ElseIf StructureName = Dom_StructureTypeEnum.SpecialStream Then
        Resolve = "SpecialStreamCatalog.csv"
    ElseIf StructureName = Dom_StructureTypeEnum.Subject Then
        Resolve = "SubjectCatalog.csv"
    ElseIf StructureName = Dom_StructureTypeEnum.UpperValues Then
        Resolve = "UpperValue.csv"
    End If
End Function
