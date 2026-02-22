Attribute VB_Name = "App_Bootstrap"
'@Folder "Application.CompositionRoot"
'ÉåÉrÉÖÅ[äœì_: ê≥ìT + ëwê”ñ±
Option Explicit

Public Sub Run()
    ' 1. Resolve Path
    Dim TopFolderPath As String
    TopFolderPath = BuiledTopFolderPath(ThisWorkbook.Path)
    Dim FilePathResolver As App_FilePathResolver
    Set FilePathResolver = New App_FilePathResolver
    FilePathResolver.Initialize TopFolderPath
    ' 2. Logger
    Dim Logger As Inf_ILogger
    Set Logger = CreateLogger(FilePathResolver)
    ' 3. INI
    Dim INIPath As String
    INIPath = BuiledINIPath(FilePathResolver)
    Dim INILoader As Inf_INIRawLoader
    Set INILoader = New Inf_INIRawLoader
    INILoader.Initialize INIPath
    Dim INISettings As Scripting.Dictionary
    Set INISettings = INILoader.ReadAll
    ' 4. Config
    ' 5. Repository
    ' 6. UseCase
    
    ' 7. Presenter
    Dim Presenter As App_Presenter
    Set Presenter = New App_Presenter
    Presenter.Initialize Logger
    ' 8. View
    Dim MainView As App_MainView
    Set MainView = New App_MainView
    Presenter.AttachView MainView
    MainView.Show
End Sub

Private Function BuiledTopFolderPath(ByVal TargetPath As String) As String
    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    BuiledTopFolderPath = FSO.GetParentFolderName(FSO.GetParentFolderName(TargetPath))
End Function

Private Function CreateLogger(ByVal Resolver As App_FilePathResolver) As Inf_ILogger
    Dim LoggerResolver As App_LoggerResolver
    Set LoggerResolver = New App_LoggerResolver
    Set CreateLogger = LoggerResolver.Resolve(Util_Environment.GetEnvironment, Resolver.ErrorlogFile)
End Function

Private Function BuiledINIPath(ByVal Resolver As App_FilePathResolver) As String
    BuiledINIPath = Resolver.DesignFile
End Function

