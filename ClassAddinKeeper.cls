VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassAddinKeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Addin")
Option Explicit
Private Directory       As IDirectory
Private UpdateVersion_  As String
Private Const DEFINED_VERSION           As String = "1.2.1"
Private Const FOLDER_PREFIX             As String = "ExBlackBoard-"
Private Const ADDIN_NAME                As String = "system"
Private Const LATEST_BOOTBOOK_NAME      As String = "LatestExBlackBoard"
Private Const UPDATE_BOOTBOOK_NAME      As String = "ExBlackBoard"
Private Const GITHUB_API                As String = "https://api.github.com/repos/LeftyBear/ExBlackBoard/releases/latest"
Private Const GITHUB_URL                As String = "https://github.com/LeftyBear/ExBlackBoard/releases/latest/"
Private Const MODULE_NAME               As String = "ClassAddinKeeper"
Private Const SHOULD_SWAP_BOOT_BOOK     As Boolean = False
Private Property Let UpdateVersion(ByVal VersionName As String)
    UpdateVersion_ = VersionName
End Property
Private Property Get UpdateVersion() As String
    UpdateVersion = UpdateVersion_
End Property
Private Sub Class_Initialize()
    UpdateVersion = FetchVersion
End Sub
Public Sub OnClickedUpdateButton()
    On Error GoTo ThrowError
    Const ProcedureName As String = "OnClickedUpdateButton"
    Const MsgTitle As String = "アップデート"
    If Not CanUpdate Then
        If UpdateVersion Like vbNullString Then
            MsgBox "アップデート情報を取得できませんでした。", vbExclamation, MsgTitle
        Else
            MsgBox "現在のバージョンは最新です。更新の必要はありません。", vbInformation, MsgTitle
        End If
        Exit Sub
    End If
    FormExBB.VerifyAddinVersion
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub OnClickedDowndateButton()
    On Error GoTo ThrowError
    Const ProcedureName As String = "OnClickedDowndateButton"
    Const MsgTitle As String = "ダウンデート"
    Dim DowndateVersions As Dictionary
    Set DowndateVersions = AddDowndateVersion
    If DowndateVersions.Count = 0 Then
        MsgBox "ダウンデートできるプログラムはありません。", vbInformation, MsgTitle: Exit Sub
    End If
    Dim BootSheetHoldsItem As IBootSheetHoldsItem
    Set BootSheetHoldsItem = New BootSheetHoldsItem
    Dim KeyVersion As Variant
    For Each KeyVersion In DowndateVersions
        If KeyVersion < BootSheetHoldsItem.AddinVersion Then
            Dim DowndateVersion As String
            DowndateVersion = DowndateVersions.Item(KeyVersion): Exit For
        End If
    Next
    If DowndateVersion Like vbNullString Then
        MsgBox "ダウンデートできるプログラムはありません。", vbInformation, MsgTitle: Exit Sub
    End If
    Dim RES As Long
    RES = MsgBox("バージョン" & BootSheetHoldsItem.AddinVersion & "から" & _
        Replace(DowndateVersion, FOLDER_PREFIX, vbNullString) & "にダウンデートしますか？", _
        vbYesNo + vbQuestion + vbDefaultButton2, MsgTitle)
    If RES <> vbYes Then Exit Sub
    Dim FolderPath As String
    Dim FilePath As String
    Set Directory = New Directory
    FolderPath = FSO.BuildPath(Directory.srcFolder, DowndateVersion)
    FilePath = FSO.BuildPath(FolderPath, BootSheetHoldsItem.AddinName)
    FSO.CopyFile FilePath, Directory.srcFolder, True
    BootSheetHoldsItem.AddinVersion = DowndateVersion
    MsgBox "バージョン" & DowndateVersion & "にダウンデートしました。", vbInformation, MsgTitle
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub TryUpdate(ByRef Result As Long)
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryUpdate"
    Const MsgTitle As String = "アップデート"
    Dim RES As Long
    RES = MsgBox("更新プログラムが見つかりました。" & vbCrLf & _
        "ダウンロードしますか？" & vbCrLf & _
        "同バージョンのアップデートが直前に失敗していた場合も再ダウンロードが必要です。", _
        vbYesNo + vbQuestion, MsgTitle)
    Select Case RES
        Case vbYes
            MsgBox "ブラウザを開きます。" & vbCrLf & _
                "更新プログラムをダウンロードしたあとはブラウザを閉じてください。", _
                vbInformation, MsgTitle
            OpenWebPageForDownload
        Case Else
            Result = vbCancel: Exit Sub
    End Select
    Application.Cursor = xlWait
    TryCopyAddinFile Result
    Application.Cursor = xlDefault
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub DeleteErrorLog()
    On Error GoTo ThrowError
    Const ProcedureName As String = "DeleteErrorLog"
    Dim Directory As IDirectory
    Set Directory = New Directory
    Dim ErrorLogPath As String
    ErrorLogPath = Directory.ErrorlogFile
    If Not FSO.FileExists(ErrorLogPath) Then Exit Sub
    FSO.DeleteFile ErrorLogPath, True
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub TrySwapBootBook()
    On Error GoTo ThrowError
    Const ProcedureName As String = "TrySwapBootBook"
    Dim UpdateBootBookFilePath As String
    Set Directory = New Directory
    UpdateBootBookFilePath = FSO.BuildPath(Directory.topFolder, UPDATE_BOOTBOOK_NAME & DEFINED_VERSION & ".xlsm")
    Application.OnTime Now + TimeValue("00:00:02"), "'OnTimeSwapBootBook " & """" & UpdateBootBookFilePath & """" & "'"
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub ShowUpdateMessage()
    On Error GoTo ThrowError
    Const ProcedureName As String = "ShowUpdateMessage"
    Dim Message As String
    Message = ReturnMessage
    Dim Note(2) As String
    Note(0) = "バージョン" & DEFINED_VERSION & "に更新されました。"
    Note(1) = vbCrLf
    Note(2) = Message
    MsgBox Join(Note, vbCrLf), vbInformation, "アップデート情報"
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub OpenWebPageForDownload()
    On Error GoTo ThrowError
    Const ProcedureName As String = "OpenWebPageForDownload"
    Dim ShellObject As IWshRuntimeLibrary.WshShell
    Set ShellObject = New IWshRuntimeLibrary.WshShell
    ShellObject.Run GITHUB_URL, 1, True
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub TryCopyAddinFile(ByRef Result As Long)
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryCopyAddinFile"
    If Not IsValidSystemFolerPath Then Result = vbAbort: Exit Sub
    Dim FileDialogTitle As String
    FileDialogTitle = FOLDER_PREFIX & UpdateVersion & ".zip"
    Dim ZipFilePath As String
    ZipFilePath = ResponsedFilePath(FileDialogTitle)
    If ZipFilePath Like vbNullString Then Result = vbCancel: Exit Sub
    If Not IsValidZipFile(ZipFilePath) Then Result = vbAbort: Exit Sub
    DeleteUnnecessaryZipFolder ZipFilePath
    Dim CurrentFolderPath As String
    CurrentFolderPath = Directory.srcFolder
    ExpandDownloadZipFile ZipFilePath, CurrentFolderPath, Result
    If Result = vbAbort Then Exit Sub
    Dim LatestFolderPath As String
    Dim LatestFilePath As String
    LatestFolderPath = FSO.BuildPath(CurrentFolderPath, FSO.GetBaseName(ZipFilePath)) & AS_FOLDER
    LatestFilePath = FSO.BuildPath(LatestFolderPath, ADDIN_NAME & ".xlam")
    FSO.CopyFile LatestFilePath, CurrentFolderPath, True
    Dim LatestAddinFilePath As String
    Dim UpdateAddinFilePath As String
    LatestAddinFilePath = FSO.BuildPath(LatestFolderPath, ADDIN_NAME & ".xlam")
    UpdateAddinFilePath = FSO.BuildPath(CurrentFolderPath, ADDIN_NAME & ".xlam")
    FSO.CopyFile LatestAddinFilePath, UpdateAddinFilePath, True
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub DeleteUnnecessaryZipFolder(ByVal ZipFilePath As String)
    On Error GoTo ThrowError
    Const ProcedureName As String = "DeleteUnnecessaryZipFolder"
    Dim Target As File
    For Each Target In FSO.GetFile(ZipFilePath).ParentFolder.Files
        If Target.Name Like FOLDER_PREFIX & "*" Then
            Target.Delete True
        End If
    Next
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub TryBuildSystemFile()
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryBuildSystemFile"
    Dim CurrentFolderPath As String
    CurrentFolderPath = Directory.srcFolder
    Dim LatestFolderPath As String
    LatestFolderPath = FSO.BuildPath(CurrentFolderPath, FOLDER_PREFIX & DEFINED_VERSION) & AS_FOLDER
    TryCopyAndBuildBootBook LatestFolderPath
    TryCopyOtherLatestFile LatestFolderPath
    TryRenameOldFolder
    TryRenameOldFile
    TryMoveOldFile
    TryDeleteOldFile
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub ExpandDownloadZipFile(ByVal DownloadFilePath As String, ByVal CurrentFolderPath As String, _
ByRef Result As Long)
    On Error GoTo ThrowError
    Const ProcedureName As String = "ExpandDownloadZipFile"
    Dim ExpandCommandlet As String
    ExpandCommandlet = LIB.JoinShellCommandForExpand(DownloadFilePath, CurrentFolderPath)
    Dim IsFailed As Boolean
    LIB.ExecuteShellCommand ExpandCommandlet, IsFailed
    If IsFailed Then Result = vbAbort: Exit Sub
    Dim DownloadFolderPath As String
    DownloadFolderPath = FSO.BuildPath(CurrentFolderPath, FSO.GetBaseName(DownloadFilePath)) & AS_FOLDER
    Dim TemporaryFilePath As String
    TemporaryFilePath = FSO.BuildPath(CurrentFolderPath, "Tmp.zip")
    Dim CompressCommandlet As String
    CompressCommandlet = LIB.JoinShellCommandForCompress(DownloadFolderPath, TemporaryFilePath)
    LIB.ExecuteShellCommand CompressCommandlet, IsFailed
    If IsFailed Then Result = vbAbort: Exit Sub
    ExpandCommandlet = LIB.JoinShellCommandForExpand(TemporaryFilePath, CurrentFolderPath)
    LIB.ExecuteShellCommand ExpandCommandlet, IsFailed
    If IsFailed Then Result = vbAbort: Exit Sub
    FSO.DeleteFile TemporaryFilePath, True
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub TryCopyAndBuildBootBook(ByVal LatestFolderPath As String)
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryCopyAndBuildBootBook"
    Dim LatestBootBookFilePath As String
    LatestBootBookFilePath = FSO.BuildPath(LatestFolderPath, LATEST_BOOTBOOK_NAME & ".xlsm")
    If Not FSO.FileExists(LatestBootBookFilePath) Then LIB.FileErrorRise LatestBootBookFilePath
    Dim UpdateBootBookFileName As String
    UpdateBootBookFileName = UPDATE_BOOTBOOK_NAME & DEFINED_VERSION & ".xlsm"
    Dim UpdateBootBookFilePath As String
    UpdateBootBookFilePath = FSO.BuildPath(Directory.topFolder, UpdateBootBookFileName)
    FSO.CopyFile LatestBootBookFilePath, UpdateBootBookFilePath, True
    With New Excel.Application
        .EnableEvents = False
        Dim UpdateBook As Workbook
        Set UpdateBook = .Workbooks.Open(UpdateBootBookFilePath)
        DoEvents
        Dim AddinBookHoldsItem As IAddinSheetHoldsItem
        Set AddinBookHoldsItem = New AddinSheetHoldsItem
        Dim UpdateSheet As Worksheet
        Set UpdateSheet = UpdateBook.Worksheets.Item(AddinBookHoldsItem.BootSheetName)
        Dim CommonCellAddress As ICommonCellAddress
        Set CommonCellAddress = New CommonCellAddress
        UpdateSheet.Range(CommonCellAddress.ForTopFolderPath).Value = Directory.topFolder
        UpdateBook.Save
        UpdateBook.Close False
        DoEvents
        .EnableEvents = True
        .Quit
    End With
    Dim Target As File
    For Each Target In FSO.GetFolder(Directory.topFolder).Files
        If Not Target.Name Like UpdateBootBookFileName Then Target.Delete True
    Next
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub TryCopyOtherLatestFile(ByVal LatestFolderPath As String)
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryCopyOtherLatestFile"
    Dim HeaderFilePath As String
    HeaderFilePath = FSO.BuildPath(LatestFolderPath, "header.csv")
    If Not FSO.FileExists(HeaderFilePath) Then LIB.FileErrorRise HeaderFilePath
    Set Directory = New Directory
    Dim SchedulesFolderPath As String
    SchedulesFolderPath = Directory.schedulesFolder
    LIB.CreateFolder SchedulesFolderPath
    FSO.CopyFile HeaderFilePath, SchedulesFolderPath, True
    Dim HolidayFilePath As String
    HolidayFilePath = FSO.BuildPath(LatestFolderPath, "syukujitsu.csv")
    If Not FSO.FileExists(HolidayFilePath) Then LIB.FileErrorRise HolidayFilePath
    Dim HolidayFolderPath As String
    HolidayFolderPath = Directory.holidayFolder
    LIB.CreateFolder HolidayFolderPath
    FSO.CopyFile HolidayFilePath, HolidayFolderPath, True
    Dim DailyFilePath As String
    DailyFilePath = FSO.BuildPath(LatestFolderPath, "Daily.csv")
    Dim WeeklyFilePath As String
    WeeklyFilePath = FSO.BuildPath(LatestFolderPath, "Weekly.csv")
    Dim MonthlyFilePath As String
    MonthlyFilePath = FSO.BuildPath(LatestFolderPath, "Monthly.csv")
    Dim YearlyFilePath As String
    YearlyFilePath = FSO.BuildPath(LatestFolderPath, "Yearly.csv")
    If Not FSO.FileExists(DailyFilePath) Then LIB.FileErrorRise DailyFilePath
    If Not FSO.FileExists(WeeklyFilePath) Then LIB.FileErrorRise WeeklyFilePath
    If Not FSO.FileExists(MonthlyFilePath) Then LIB.FileErrorRise MonthlyFilePath
    If Not FSO.FileExists(YearlyFilePath) Then LIB.FileErrorRise YearlyFilePath
    Dim TemplatesFolderPath As String
    TemplatesFolderPath = Directory.templatesFolder
    LIB.CreateFolder TemplatesFolderPath
    FSO.CopyFile DailyFilePath, TemplatesFolderPath, True
    FSO.CopyFile WeeklyFilePath, TemplatesFolderPath, True
    FSO.CopyFile MonthlyFilePath, TemplatesFolderPath, True
    FSO.CopyFile YearlyFilePath, TemplatesFolderPath, True
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub TryRenameOldFolder()
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryRenameOldFolder"
    Dim ScheduleFolderPath As String
    ScheduleFolderPath = FSO.BuildPath(Directory.dataFolder, "schedule") & AS_FOLDER
    If FSO.FolderExists(ScheduleFolderPath) Then
        FSO.GetFolder(ScheduleFolderPath).Name = "schedules"
    End If
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub TryRenameOldFile()
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryRenameOldFile"
    Dim RegulardayFilePath As String
    Set Directory = New Directory
    RegulardayFilePath = FSO.BuildPath(Directory.regularFolder, "regularday.csv")
    If FSO.FileExists(RegulardayFilePath) Then
        FSO.GetFile(RegulardayFilePath).Name = "RegularWeekDays.csv"
    End If
    Dim RegulardateFilePath As String
    RegulardateFilePath = FSO.BuildPath(Directory.regularFolder, "regulardate.csv")
    If FSO.FileExists(RegulardateFilePath) Then
        FSO.GetFile(RegulardateFilePath).Name = "RegularDays.csv"
    End If
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub TryMoveOldFile()
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryMoveOldFile"
    Dim OldDaydutyFilePath As String
    Set Directory = New Directory
    OldDaydutyFilePath = FSO.BuildPath(Directory.dataFolder, "dayduty.csv")
    Dim NewDaydutyFilePath As String
    NewDaydutyFilePath = Directory.DayDutyFile
    If Not FSO.FileExists(NewDaydutyFilePath) Then
        If FSO.FileExists(OldDaydutyFilePath) Then
            LIB.CreateFolder Directory.daydutyFolder
            FSO.MoveFile OldDaydutyFilePath, NewDaydutyFilePath
        End If
    End If
    Dim OldSyukujitsuFilePath As String
    OldSyukujitsuFilePath = FSO.BuildPath(Directory.dataFolder, "syukujitsu.csv")
    Dim NewHolidayFilePath As String
    NewHolidayFilePath = Directory.HolidayFile
    If Not FSO.FileExists(NewHolidayFilePath) Then
        If FSO.FileExists(OldSyukujitsuFilePath) Then
            LIB.CreateFolder Directory.holidayFolder
            FSO.MoveFile OldSyukujitsuFilePath, NewHolidayFilePath
        End If
    End If
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub TryDeleteOldFile()
    On Error GoTo ThrowError
    Const ProcedureName As String = "TryDeleteOldFile"
    Dim OldErrorLogPath As String
    OldErrorLogPath = FSO.BuildPath(Directory.dataFolder, "error.log")
    If FSO.FileExists(OldErrorLogPath) Then FSO.DeleteFile OldErrorLogPath, True
Exit Sub
ThrowError:
    ErrorHandler.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Function DefindVersion() As String
    DefindVersion = DEFINED_VERSION
End Function
Public Function CanUpdate() As Boolean
    If UpdateVersion Like vbNullString Then Exit Function
    If UpdateVersion <= DEFINED_VERSION Then Exit Function
    CanUpdate = True
End Function
Public Function IsLatestVersion() As Boolean
    Dim BootSheetHoldsItem As IBootSheetHoldsItem
    Set BootSheetHoldsItem = New BootSheetHoldsItem
    If BootSheetHoldsItem.AddinVersion Like DEFINED_VERSION Then IsLatestVersion = True
End Function
Public Function ShouldSwapBootBook() As Boolean
    If Not SHOULD_SWAP_BOOT_BOOK Then Exit Function
    ShouldSwapBootBook = True
End Function
Private Function ResponsedFilePath(ByVal DownloadFileName As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "更新プログラム［" & DownloadFileName & "］の選択"
        .AllowMultiSelect = False
        If .Show = False Then Exit Function
        ResponsedFilePath = .SelectedItems.Item(1)
    End With
End Function
Private Function AddDowndateVersion() As Dictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim Target As Folder
    For Each Target In FSO.GetFolder(Directory.srcFolder).SubFolders
        RET.Add Replace(Target.Name, FOLDER_PREFIX, vbNullString), Target.Name
    Next
    If RET.Count = 0 Then Exit Function
    Set AddDowndateVersion = LIB.SortDictionary(RET)
End Function
Private Function ReturnMessage() As String
    Dim Note(4) As String
    Note(0) = "◆データが正常に保存されなかった場合の対策を施しました。"
    Note(1) = "◆行事欄を上下２つに分け、上段の行事のみ年間行事予定に印字するように変更しました。"
    Note(2) = "◆教育相談と来校を日誌に印字するように変更しました。"
    Note(3) = "◆黒板を最小化したあとの再表示は「Ex黒板」のロゴをクリックする仕様に変更しました。"
    Note(4) = "◆それらのほかにも様々な修正を行っています。"
    ReturnMessage = Join(Note, vbCrLf)
End Function
Private Function FetchVersion() As String
    Dim HttpObject As XMLHTTP60
    Set HttpObject = New XMLHTTP60
    HttpObject.Open "GET", GITHUB_API, False
    HttpObject.send
    Do While HttpObject.readyState < 4
        DoEvents
    Loop
    Dim TextString As String
    TextString = HttpObject.responseText()
    If Not TextString Like "*tag_name*" Then Exit Function
    Dim RET As String
    RET = Mid$(TextString, InStr(TextString, "tag_name") + Len("""tag_name""") + 1)
    RET = LIB.MidFromLeftToBeforeTargetChar(RET, CANNMA)
    RET = LIB.ErasureFromLeft(RET, 1)
    RET = Left$(RET, Len(RET) - 1)
    FetchVersion = RET
End Function
Private Function IsValidZipFile(ByVal FilePath As String) As Boolean
    Const MsgTitle As String = "アップデート"
    IsValidZipFile = True
    If FSO.GetExtensionName(FilePath) Like "zip" Then
        If FSO.GetBaseName(FilePath) Like FOLDER_PREFIX & UpdateVersion Then Exit Function
    End If
    IsValidZipFile = False
    MsgBox "無効なファイルを選択しました。" & vbCrLf & "処理を中断します。", vbExclamation, MsgTitle
End Function
Private Function IsValidSystemFolerPath() As Boolean
    Dim FolderNames() As String
    FolderNames = Split(Directory.topFolder, AS_FOLDER)
    Dim HasSpace() As String
    HasSpace = Filter(FolderNames, HALF_SPACE, compare:=vbTextCompare)
    If UBound(HasSpace) > -1 Then
        MsgBox "フォルダ名にスペースを含めることはできません。" & vbCrLf & _
            "フォルダ名を変更してから再起動してください。" & vbCrLf & _
            "対象：" & Join(HasSpace, "、"), vbExclamation, "無効なフォルダ名"
        Exit Function
    End If
    IsValidSystemFolerPath = True
End Function
'@Ignore ProcedureNotUsed
Public Property Let IsSwappingBootBook(ByVal FilePath As String, ByVal State As Boolean)
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        .WriteLine CStr(State)
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Public Property Get IsSwappingBootBook(ByVal FilePath As String) As Boolean
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForReading)
        IsSwappingBootBook = CBool(.ReadLine)
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property


