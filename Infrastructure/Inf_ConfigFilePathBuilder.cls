VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ConfigFilePathBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service")
Option Explicit

Private Type Member
    Resolver As Inf_ConfigFileNameResolver
    Provider As Inf_WorkbookPathProvider
End Type

Private This As Member

Public Sub Initialize(ByVal Resolver As Inf_ConfigFileNameResolver, ByVal Provider As Inf_WorkbookPathProvider)
    Set This.Resolver = Resolver
    Set This.Provider = Provider
End Sub

Public Function Builed(ByVal ConfigName As Dom_Config) As String
    Dim FileName As String
    FileName = This.Resolver.Resolve(ConfigName)
    Builed = This.Provider.GetBaseFolderPath & "root\config\" & FileName
End Function
