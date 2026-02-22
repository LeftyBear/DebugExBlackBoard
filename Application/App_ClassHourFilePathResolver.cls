VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourFilePathResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Application.Resolver
Option Explicit

Private Type Member
    YearFolderResolver As App_YearFolderResolver
    FileNameResolver As App_FileNameResolver
End Type

Private This As Member

Public Sub Initialize(ByVal YearFolderResolver As App_YearFolderResolver, ByVal FileNameResolver As App_FileNameResolver)
    Set This.YearFolderResolver = YearFolderResolver
    Set This.FileNameResolver = FileNameResolver
End Sub

Public Function Resolve(ByVal SchoolYearKey As App_SchoolYearKey) As String
    Dim Provider As Inf_WorkbookPathProvider
    Set Provider = New Inf_WorkbookPathProvider
    Dim BasePathResolver As App_BasePathResolver
    Set BasePathResolver = New App_BasePathResolver
    BasePathResolver.Initialize Provider
    Dim BasePath As String
    BasePath = BasePathResolver.Resolve
    Dim FolderPath As String
    FolderPath = This.YearFolderResolver.Resolve(BasePath, SchoolYearKey)
    Dim FileName As String
    FileName = This.FileNameResolver.Resolve(SchoolYearKey, ClassHour)
    Resolve = FolderPath & "\" & FileName
End Function


