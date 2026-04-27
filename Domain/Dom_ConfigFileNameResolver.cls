VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ConfigFileNameResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Resolver")
Option Explicit

Public Function Resolve(ByVal StructureName As Dom_StructureTypeEnum) As String
    Select Case StructureName
    Case Dom_StructureTypeEnum.SchoolEvent
        Resolve = "SchoolEventCatalog.csv"
    Case Dom_StructureTypeEnum.SpecialStream
        Resolve = "SpecialStreamCatalog.csv"
    Case Dom_StructureTypeEnum.Subject
        Resolve = "SubjectCatalog.csv"
    Case Dom_StructureTypeEnum.UpperValues
        Resolve = "UpperValue.csv"
    End Select
End Function
