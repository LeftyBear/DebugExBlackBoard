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

Public Function Resolve(ByVal StructureName As Dom_StructureName) As String
    Select Case StructureName
    Case Dom_StructureName.UpperValues
        Resolve = "UpperValue.csv"
    Case Dom_StructureName.SpecialStream
        Resolve = "SpecialStreamCatalog.csv"
    Case Dom_StructureName.Subject
        Resolve = "SubjectCatalog.csv"
    End Select
End Function

