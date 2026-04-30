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

Public Function Resolve(ByVal TypeCode As Dom_StructureTypePolicy) As String
    If TypeCode = Dom_StructureTypePolicy.SchoolEvent Then
        Resolve = "SchoolEventCatalog.csv"
    ElseIf TypeCode = Dom_StructureTypePolicy.SpecialStream Then
        Resolve = "SpecialStreamCatalog.csv"
    ElseIf TypeCode = Dom_StructureTypePolicy.Subject Then
        Resolve = "SubjectCatalog.csv"
    ElseIf TypeCode = Dom_StructureTypePolicy.UpperValues Then
        Resolve = "UpperValue.csv"
    End If
End Function
