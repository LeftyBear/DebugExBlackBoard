VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ConfigFilePathBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Resolver As Dom_ConfigFileNameResolver
    Provider As App_IWorkbookPathProvider
End Type

Private This As Member

Public Sub Initialize(ByVal Resolver As Dom_ConfigFileNameResolver, ByVal Provider As App_IWorkbookPathProvider)
    Set This.Resolver = Resolver
    Set This.Provider = Provider
End Sub

Public Function Builed(ByVal StructureName As Dom_StructureTypeEnum) As String
    Dim FileName As String
    FileName = This.Resolver.Resolve(StructureName)
    Builed = This.Provider.GetBaseFolderPath & "root\Config\" & FileName
End Function
