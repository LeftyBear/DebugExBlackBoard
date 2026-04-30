VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EntityFilePathBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function Builed(ByVal TypeCode As Dom_EntityTypePolicy, ByVal SchoolYear As Long) As String
    Dim Provider As Inf_WorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim Resolver As Dom_EntityFileNameResolver
    Set Resolver = New Dom_EntityFileNameResolver
    Dim FileName As String
    FileName = Resolver.Resolve(TypeCode, SchoolYear)
    Builed = Inf_StringUtility.BuildToPath(Provider.GetBaseFolderPath, "root", "data", "schoolyears", SchoolYear, FileName)
End Function
