VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Backup")
Option Explicit
Private FileNames_          As Collection
Private LogFolderPath       As String
Private BackupFolderPath    As String
Private Const Incomplete    As String = "Incomplete"
Private Const Complete      As String = "Complete"
Public Property Set FileNames(ByVal Target As Collection)
    Set FileNames_ = Target
End Property
Public Property Get FileNames() As Collection
    Set FileNames = FileNames_
End Property
Private Sub Class_Initialize()
    LIB.CreateFolder System.Directory.writelogFolder
    LIB.CreateFolder System.Directory.backupFolder
End Sub
Public Sub WriteFileAfterBackup(ByVal FilePath As String, ByRef DataBody() As Variant)
    If FSO.FileExists(FilePath) Then
        CopyFile FilePath
        WriteIncompleteLog FilePath
    End If
    LIB.WriteSomeLines FilePath, DataBody
    WriteCompleteLog FilePath
End Sub
Private Sub WriteIncompleteLog(ByVal FilePath As String)
    Dim LogName As String
    Dim LogPath As String
    LogName = FSO.GetBaseName(FilePath) & ".log"
    LogPath = FSO.BuildPath(LogFolderPath, LogName)
    Dim Notes(1) As String
    Dim Text As String
    Notes(0) = Incomplete
    Notes(1) = FilePath
    Text = Join(Notes, CANNMA)
    LIB.WriteOneLine LogPath, Text
End Sub
Private Sub WriteCompleteLog(ByVal FilePath As String)
    Dim LogName As String
    Dim LogPath As String
    LogName = FSO.GetBaseName(FilePath) & ".log"
    LogPath = FSO.BuildPath(LogFolderPath, LogName)
    Dim Notes(1) As String
    Dim Text As String
    Notes(0) = Complete
    Notes(1) = FilePath
    Text = Join(Notes, CANNMA)
    LIB.WriteOneLine LogPath, Text
End Sub
Private Sub CopyFile(ByVal FilePath As String)
    FSO.CopyFile FilePath, BackupFolderPath, True
End Sub
Public Sub RestoreFile()
    Dim FileName As Variant
    For Each FileName In Me.FileNames
        TryRestore CStr(FileName)
    Next
End Sub
Private Sub TryRestore(ByVal FileName As String)
    Dim BackupFile As File
    For Each BackupFile In FSO.GetFolder(BackupFolderPath).Files
        If FSO.GetBaseName(BackupFile.Name) Like FileName Then
            Dim LogFilePath As String
            Dim OriginalFilePath As String
            LogFilePath = FSO.BuildPath(LogFolderPath, FileName & ".log")
            OriginalFilePath = LIB.ReadOneLineAsArray1D(LogFilePath)(1)
            FSO.CopyFile BackupFile.Path, OriginalFilePath, True
        End If
    Next
End Sub
Public Function IsNothing() As Boolean
    Dim BUF As Collection
    Set BUF = New Collection
    Dim Target As File
    For Each Target In FSO.GetFolder(LogFolderPath).Files
        Dim State As String
        State = LIB.ReadAllLine(Target.Path)
        If State Like Incomplete & "*" Then BUF.Add FSO.GetBaseName(Target.Name)
    Next
    Set Me.FileNames = BUF
    If BUF.Count = 0 Then IsNothing = True
End Function

