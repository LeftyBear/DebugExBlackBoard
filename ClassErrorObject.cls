VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassErrorObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ErrorObject")
Option Explicit
Private EventName_                  As String
Private ModuleName_                 As String
Private ProcedureName_              As String
Private Const MODULE_NAME           As String = "MODULE_NAME"
Private Const EVENT_NAME            As String = "EVENT_NAME"
Private Const PROCEDURE_NAME        As String = "PROCEDURE_NAME"
Private Const ERROR_NUMBER          As String = "ERROR_NUMBER"
Private Const ERROR_DESCRIPTION     As String = "ERROR_DESCRIPTION"
Private Const MASSAGE_TEMPLATE      As String = _
    "予期せぬエラーが発生したので処理を中断しました。" & vbCrLf & _
    "詳しくはシステム管理者にお問い合わせください。" & vbCrLf & _
    "モジュール名" & vbTab & MODULE_NAME & vbCrLf & _
    "イベント名" & vbTab & vbTab & EVENT_NAME & vbCrLf & _
    "プロシージャ名" & vbTab & PROCEDURE_NAME & vbCrLf & _
    "エラー番号" & vbTab & vbTab & ERROR_NUMBER & vbCrLf & _
    "エラー説明" & vbTab & vbTab & ERROR_DESCRIPTION
Public Sub CatchEvent(ByVal EventName As String)
    If Not EventName_ Like vbNullString Then Exit Sub
    EventName_ = EventName
End Sub
Public Sub Throw(ByVal ModuleName As String, ByVal ProcedureName As String)
    If Not ProcedureName Like vbNullString Then Exit Sub
    ModuleName_ = ModuleName
    ProcedureName_ = ProcedureName
End Sub
Public Sub ShowErrorMessage()
    Dim errorlog As String
    errorlog = MASSAGE_TEMPLATE
    errorlog = Replace(errorlog, MODULE_NAME, ModuleName_)
    errorlog = Replace(errorlog, EVENT_NAME, EventName_)
    errorlog = Replace(errorlog, PROCEDURE_NAME, ProcedureName_)
    errorlog = Replace(errorlog, ERROR_NUMBER, Err.Number)
    errorlog = Replace(errorlog, ERROR_DESCRIPTION, Err.Description)
    MsgBox errorlog, vbExclamation, "予期せぬエラー"
    Dim ErrorLogs() As Variant
    ErrorLogs = Array(Now, ModuleName_, EventName_, ProcedureName_, Err.Number, Err.Description)
    WriteErrorLog Join(ErrorLogs, CANNMA)
    Me.ClearCache
    LIB.Application画面更新 = True
End Sub
Public Sub ClearCache()
    EventName_ = vbNullString
    ModuleName_ = vbNullString
    ProcedureName_ = vbNullString
End Sub
Private Sub WriteErrorLog(ByVal errorlog As String)
    LIB.CreateFolder CSO.GetPath.data.FolderErrorLog
    Dim FilePath As String
    FilePath = CSO.GetPath.data.errorlog.FileErrorLog
    On Error Resume Next
    LIB.WriteOneLineForAppending FilePath, errorlog
    On Error GoTo 0
End Sub
