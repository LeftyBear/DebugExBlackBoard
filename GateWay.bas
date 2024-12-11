Attribute VB_Name = "GateWay"
'@Folder("Addin")
Option Explicit
Option Private Module
'@Ignore EncapsulatePublicField
Public CSO      As New ClassSystemObject
'@Ignore EncapsulatePublicField
Public CEO      As New ClassErrorObject
'@Ignore EncapsulatePublicField
Public Admin    As New ClassAdministrator
Private Const MODULE_NAME   As String = "Gateway"
Private Sub RefreshAddin()
    On Error GoTo ThrowError
    Const ProcedureName As String = "RefreshAddin"
    With CSO
        .BootBookName = vbNullString
        .BootSheetName = vbNullString
    End With
    Schedule.DeleteTable
    Enrollment.DeleteTable
    ClassHour.DeleteTable
    Addinbook.IsAddin = True
    Addinbook.Save
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
'@EntryPoint
Private Sub ShowUserForm()
    CEO.CatchEvent "ShowUserForm"
    On Error GoTo Finaly
    SetupFormExBB
    ShowFormExBB
    CEO.ClearCache
Exit Sub
Finaly:
    CEO.ShowErrorMessage
End Sub
'@EntryPoint
Private Sub ConnectWithAddin(ByVal BookName As String, ByVal SheetName As String)
    CEO.CatchEvent "ConnectWithAddin"
    On Error GoTo Finaly
    ResetSystemInformation BookName, SheetName
    If Not IsValidBootBook Then DisconnctFromAddin
    VerifyAddinVersion
    If CSO.IsSwappingBootBook Then CEO.ClearCache: Exit Sub
    SetupFormExBB
    ShowFormExBB
    ProtectBookAndSheet
    BackUpData
    CEO.ClearCache
Exit Sub
Finaly:
    CEO.ShowErrorMessage
End Sub
Public Sub VerifyAddinVersion()
    On Error GoTo ThrowError
    Const ProcedureName As String = "VerifyAddinVersion"
    Dim AddinKeeper As ClassAddinKeeper
    Set AddinKeeper = New ClassAddinKeeper
    If CSO.IsAdministrator Then VerifyUpdate AddinKeeper
    If AddinKeeper.IsLatestVersion Then Exit Sub
    With AddinKeeper
        If CSO.IsAdministrator Then .TryBuildSystemFile
'        If Not CSO.IsSwappingBootBook Then
'            .TrySwapBootBook
'            CSO.IsSwappingBootBook = True
'            Exit Sub
'        End If
'        CSO.IsSwappingBootBook = False
        .ShowUpdateMessage
        CSO.AddinVersion = .DefindVersion
    End With
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub VerifyUpdate(ByVal AddinKeeper As ClassAddinKeeper)
    On Error GoTo ThrowError
    Const ProcedureName As String = "VerifyUpdate"
    Const MsgTitle As String = "アップデート"
    If Not AddinKeeper.CanUpdate Then Exit Sub
    Dim Result As Long
    AddinKeeper.TryUpdate Result
    Select Case Result
        Case vbEmpty
            AddinKeeper.DeleteErrorLog
            MsgBox "更新プログラムのインストールが完了しました。" & vbCrLf & _
                "システムを再起動してください。", vbInformation, MsgTitle
            DisconnctFromAddin
        Case vbCancel
            MsgBox "更新プログラムのダウンロードを中断しました。", vbExclamation, MsgTitle: Exit Sub
        Case vbAbort
            MsgBox "更新プログラムのダウンロードに失敗しました。", vbExclamation, MsgTitle: Exit Sub
    End Select
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
'@EntryPoint
Public Sub OnTimeSwapBootBook(ByVal UpdateBootBookFilePath As String)
    On Error GoTo ThrowError
    Const ProcedureName As String = "OnTimeSwapBootBook"
    Application.Cursor = xlWait
    Dim CurrentBootBookFilePath As String
    CurrentBootBookFilePath = CSO.BootBook.FullName
    Dim CurrentBootBookFolderPath As String
    Dim UpdateBootBookFileName As String
    CurrentBootBookFolderPath = CSO.BootBook.Path
    UpdateBootBookFileName = FSO.GetFileName(UpdateBootBookFilePath)
    Dim LatestBootBookFilePath As String
    LatestBootBookFilePath = FSO.BuildPath(CurrentBootBookFolderPath, UpdateBootBookFileName)
    CSO.BootBook.Close False
    DoEvents
    FSO.DeleteFile CurrentBootBookFilePath, True
    DoEvents
    FSO.CopyFile UpdateBootBookFilePath, LatestBootBookFilePath, True
    DoEvents
    Application.Cursor = xlDefault
    Workbooks.Open LatestBootBookFilePath
    DoEvents
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub DisconnctFromAddin()
    Unload FormExBB
    With CSO
        .BootBook.Unprotect Password:=Admin.ProtectionPassword
        .BootSheet.Unprotect Password:=Admin.ProtectionPassword
        .BootBook.Save
        If .BootBookName Like "*debug*" Then RefreshAddin
    End With
    Dim Book As Workbook
    For Each Book In Application.Workbooks
        If Not UCase$(FSO.GetExtensionName(Book.Name)) Like UCase$("*xlsb*") Then
            Dim BookCount As Long
            BookCount = BookCount + 1
        End If
    Next
    If BookCount = 1 Then
        Application.Quit
        Addinbook.Close False
    Else
        CSO.BootBook.Close False
        Addinbook.Close False
    End If
End Sub
Private Sub ResetSystemInformation(ByVal BootBookName As String, ByVal BootSheetName As String)
    On Error GoTo ThrowError
    Const ProcedureName As String = "ResetSystemInformation"
    LIB.Application画面更新 = False
    With CSO
        .BootBookName = BootBookName
        .BootSheetName = BootSheetName
        .AddinName = Addinbook.Name
        .CurrentYear = vbNullString
        .CurrentScheduleFileUpdateAt = 0
        If .IsAdministrator Then
            .backupFolder = FSO.BuildPath(CurDir$, "ExBlackBoardBackupData") & AS_FOLDER
        End If
    End With
    LIB.Application画面更新 = True
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub SetupFormExBB()
    On Error GoTo ThrowError
    Const ProcedureName As String = "SetUpFormExBB"
    With FormExBB
        .SetUpForm
        .OnChangedDate Date
    End With
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub ShowFormExBB()
    On Error GoTo ThrowError
    Const ProcedureName As String = "ShowFormExBB"
    With New ClassFormSizing
        Set .TargetForm = FormExBB
        .EnabledMaximization = True
        .EnabledMinimization = True
        .EnabledResizing = True
        .EnabledClosingButton = False
        .DrawFormMenuBar
        .MaximizeForm
    End With
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub ProtectBookAndSheet()
    On Error GoTo ThrowError
    Const ProcedureName As String = "ProtectBookAndSheet"
    If Not CSO.BootBookName Like "*debug*" Then Exit Sub
    With CSO
        .BootBook.Protect Password:=Admin.ProtectionPassword
        .BootSheet.Protect Password:=Admin.ProtectionPassword, UserInterfaceOnly:=True
    End With
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub BackUpData()
    On Error GoTo ThrowError
    Const ProcedureName As String = "BackUpData"
    If Not CSO.IsAdministrator Then Exit Sub
    Dim DataFolderPath As String
    DataFolderPath = LIB.ErasureFromRight(CSO.GetPath.FolderData, Len(AS_FOLDER))
    If Not FSO.FolderExists(DataFolderPath) Then Exit Sub
    Dim backupFolder As String
    backupFolder = CSO.backupFolder
    LIB.CreateFolder backupFolder
    If FSO.GetFolder(backupFolder).SubFolders.Count > 5 Then
        Dim Target As Folder
        For Each Target In FSO.GetFolder(backupFolder).SubFolders
            If IsNumeric(Target.Name) Then
                Dim OldName As String
                If OldName Like vbNullString Then OldName = Target.Name
                If OldName < Target.Name Then
                    OldName = Target.Name
                    Dim TargetPath As String
                    TargetPath = Target.Path
                End If
            End If
        Next
        FSO.DeleteFolder TargetPath, True
    End If
    Dim BackupFolderPath As String
    BackupFolderPath = FSO.BuildPath(backupFolder, Format$(Date, "yyyymmdd") & Format$(Time, "hhmmss"))
    FSO.CopyFolder DataFolderPath, BackupFolderPath, True
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Function IsValidBootBook() As Boolean
    Dim BootBookPath As String
    BootBookPath = CSO.BootBook.FullName
    Dim BootBookPathOfTopFolder As String
    BootBookPathOfTopFolder = FSO.BuildPath(CSO.GetPath.FolderTop, CSO.BootBookName)
    If BootBookPath Like BootBookPathOfTopFolder Then
        MsgBox "トップフォルダ上の起動用ブックは操作できません。" & vbCrLf & _
            "トップフォルダ上の起動用ブックをデスクトップ等にコピーしてから開いてください。" & _
            "（ショートカットは使用できません。）" & vbCrLf & _
            "システムを終了します。", vbExclamation, "システム起動不可"
        Exit Function
    End If
    IsValidBootBook = True
End Function

