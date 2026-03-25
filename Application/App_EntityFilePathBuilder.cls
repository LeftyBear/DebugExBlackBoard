VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EntityFilePathBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Builder")
Option Explicit

Private Type Member
    Resolver As Dom_EntityFileNameResolver
    Provider As App_IWorkbookPathProvider
End Type

Private This As Member

Public Sub Initialize(ByVal Provider As App_IWorkbookPathProvider)
    Set This.Resolver = New Dom_EntityFileNameResolver
    Set This.Provider = Provider
End Sub

Public Function Builed(ByVal EntityName As Dom_EntityName, ByVal SchoolYear As Dom_SchoolYear) As String
    Dim FileName As String
    FileName = This.Resolver.Resolve(EntityName, SchoolYear)
    Builed = This.Provider.GetBaseFolderPath & "root\data\schoolyears\" & SchoolYear.ToKey & charBackSlash & FileName
End Function
