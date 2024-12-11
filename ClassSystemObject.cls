VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Private Const ADDRESS_ADDIN_NAME            As String = "B1"
Private Const ADDRESS_ADDIN_VERSION         As String = "B2"
Private Const ADDRESS_TOP_FOLDER_PATH       As String = "B3"
Private Const ADDRESS_BOOT_BOOK_NAME        As String = "B4"
Private Const ADDRESS_BOOT_SHEET_NAME       As String = "B5"
Private Const ADDRESS_SCHEDULE_BY_YEAR      As String = "B6"
Private Const ADDRESS_SCHEDULE_UPDATEAT     As String = "B7"
Private Const ADDRESS_BACKUP_FOLDER_PATH    As String = "B8"
Public Property Let BootBookName(ByVal BookName As String)
    AddinSheet.Range(ADDRESS_BOOT_BOOK_NAME).Value = BookName
End Property
Public Property Get BootBookName() As String
    BootBookName = AddinSheet.Range(ADDRESS_BOOT_BOOK_NAME).Value
End Property
Public Property Let BootSheetName(ByVal SheetName As String)
    AddinSheet.Range(ADDRESS_BOOT_SHEET_NAME).Value = SheetName
End Property
Public Property Get BootSheetName() As String
    BootSheetName = AddinSheet.Range(ADDRESS_BOOT_SHEET_NAME).Value
End Property
Public Property Let AddinName(ByVal TargetName As String)
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    Sheet.Range(ADDRESS_ADDIN_NAME).Value = TargetName
End Property
Public Property Get AddinName() As String
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    AddinName = Sheet.Range(ADDRESS_ADDIN_NAME).Value
End Property
Public Property Let AddinVersion(ByVal TargetVersion As String)
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    Sheet.Range(ADDRESS_ADDIN_VERSION).Value = TargetVersion
End Property
Public Property Get AddinVersion() As String
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    AddinVersion = Sheet.Range(ADDRESS_ADDIN_VERSION).Value
End Property
Public Property Let CurrentYear(ByVal WesternYear As String)
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    Sheet.Range(ADDRESS_SCHEDULE_BY_YEAR).Value = WesternYear
End Property
Public Property Get CurrentYear() As String
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    CurrentYear = Sheet.Range(ADDRESS_SCHEDULE_BY_YEAR).Value
End Property
Public Property Let CurrentScheduleFileUpdateAt(ByVal TargetDate As Date)
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    Sheet.Range(ADDRESS_SCHEDULE_UPDATEAT).Value = TargetDate
End Property
Public Property Get CurrentScheduleFileUpdateAt() As Date
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    CurrentScheduleFileUpdateAt = Sheet.Range(ADDRESS_SCHEDULE_UPDATEAT).Value
End Property
Public Property Let backupFolder(ByVal FolderPath As String)
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    Sheet.Range(ADDRESS_BACKUP_FOLDER_PATH).Value = FolderPath
End Property
Public Property Get backupFolder() As String
    Dim Sheet As Worksheet
    Set Sheet = Me.BootSheet
    If Sheet Is Nothing Then Exit Property
    backupFolder = Sheet.Range(ADDRESS_BACKUP_FOLDER_PATH).Value
End Property
Public Property Let Administrator(ByVal TargetName As String)
    On Error GoTo ThrowError
    LIB.CreateFolder Me.GetPath.data.FolderCache
    Dim FilePath As String
    FilePath = Me.GetPath.data.cache.FileAdministrator
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        .WriteLine TargetName
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Public Property Get Administrator() As String
    On Error GoTo ThrowError
    Dim FilePath As String
    FilePath = Me.GetPath.data.cache.FileAdministrator
    If Not FSO.FileExists(FilePath) Then Exit Property
    With FSO.OpenTextFile(FilePath, ForReading)
        Administrator = .ReadLine
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Public Property Let AdministrationPassword(ByVal TargetPassword As String)
    On Error GoTo ThrowError
    LIB.CreateFolder Me.GetPath.data.FolderCache
    Dim FilePath As String
    FilePath = Me.GetPath.data.cache.FilePassword
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        .WriteLine TargetPassword
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Public Property Get AdministrationPassword() As String
    On Error GoTo ThrowError
    Dim FilePath As String
    FilePath = Me.GetPath.data.cache.FilePassword
    If Not FSO.FileExists(FilePath) Then Exit Property
    With FSO.OpenTextFile(FilePath, ForReading)
        AdministrationPassword = .ReadLine
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Public Property Let SkipsPassword(ByVal State As Boolean)
    On Error GoTo ThrowError
    LIB.CreateFolder Me.GetPath.data.FolderCache
    Dim FilePath As String
    FilePath = Me.GetPath.data.cache.FilePasswordSkip
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        .WriteLine CStr(State)
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Public Property Get SkipsPassword() As Boolean
    On Error GoTo ThrowError
    Dim FilePath As String
    FilePath = Me.GetPath.data.cache.FilePasswordSkip
    If Not FSO.FileExists(FilePath) Then Exit Property
    With FSO.OpenTextFile(FilePath, ForReading)
        SkipsPassword = CBool(.ReadLine)
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
'Public Property Let IsSwappingBootBook(ByVal State As Boolean)
'    On Error GoTo ThrowError
'    LIB.CreateFolder Me.GetPath.data.FolderCache
'    Dim FilePath As String
'    FilePath = Me.GetPath.data.cache.FileOnTime
'    With FSO.OpenTextFile(FilePath, ForWriting, True)
'        .WriteLine CStr(State)
'ThrowError:
'        .Close
'        DoEvents
'    End With
'    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
'End Property
Public Property Get IsSwappingBootBook() As Boolean
    On Error GoTo ThrowError
    Dim FilePath As String
    FilePath = Me.GetPath.data.cache.FileOnTime
    If Not FSO.FileExists(FilePath) Then Exit Property
    With FSO.OpenTextFile(FilePath, ForReading)
        IsSwappingBootBook = CBool(.ReadLine)
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Public Function BootBook() As Workbook
    Dim BookName As String
    BookName = Me.BootBookName
    If BookName Like vbNullString Then Exit Function
    Set BootBook = Workbooks.Item(BookName)
End Function
Public Function BootSheet() As Worksheet
    Dim BootBook As Workbook
    Set BootBook = Me.BootBook
    If BootBook Is Nothing Then Exit Function
    Dim SheetName As String
    SheetName = Me.BootSheetName
    If SheetName Like vbNullString Then Exit Function
    Set BootSheet = BootBook.Worksheets.Item(SheetName)
End Function
Public Function AddressOfTopFolderPath() As String
    AddressOfTopFolderPath = ADDRESS_TOP_FOLDER_PATH
End Function
Public Function GetPath() As ClassPath
    Set GetPath = New ClassPath
End Function
Public Function IsAdministrator() As Boolean
    If Not Me.Administrator Like Environ$("USERNAME") Then Exit Function
    IsAdministrator = True
End Function
Public Function HasSetedPassword() As Boolean
    If Me.AdministrationPassword Like vbNullString Then
        MsgBox "パスワードが設定されていません。" & vbCrLf & _
            "パスワードを設定してください。", vbExclamation, "管理者パスワード設定"
        Exit Function
    End If
    HasSetedPassword = True
End Function


