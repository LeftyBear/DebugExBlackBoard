VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_EntityFilePathBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Service")
Option Explicit

Private Type Member
    Resolver As Inf_EntityFileNameResolver
    Provider As Inf_WorkbookPathProvider
End Type

Private This As Member

Public Sub Initialize(ByVal Resolver As Inf_EntityFileNameResolver, ByVal Provider As Inf_WorkbookPathProvider)
    Set This.Resolver = Resolver
    Set This.Provider = Provider
End Sub

Public Function Builed(ByVal EntityName As Dom_Entity, ByVal SchoolYear As Long) As String
    Dim FileName As String
    FileName = This.Resolver.Resolve(EntityName, SchoolYear)
    Builed = This.Provider.GetBaseFolderPath & "root\data\schoolyears\" & CStr(SchoolYear) & charBackSlash & FileName
End Function
