VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Administrator")
Option Explicit
Private Const SHEET_PROTECT_PASSWORD    As String = "pass3225"
Private Const ENCRYPT_KEY               As String = "f4jO1cYz"
Private Const MODULE_NAME               As String = "ClassAdministrator"
Public Sub SetPassword()
    Const ProcedureName As String = "SetPassword"
    On Error GoTo ThrowError
    Const MsgTitle As String = "管理者パスワード設定"
    Do
        Dim NewPassword As String
        NewPassword = InputBox("新しいパスワードを入力してください。", MsgTitle)
        If NewPassword Like vbNullString Then Exit Sub
        Dim ConfirmPassword As String
        ConfirmPassword = InputBox("確認のためにもう一度パスワードを入力してください。", MsgTitle)
        If ConfirmPassword Like vbNullString Then Exit Sub
        If NewPassword = ConfirmPassword Then Exit Do
        MsgBox "パスワードが一致しません。はじめからやり直してください。", vbExclamation, MsgTitle
    Loop
    CSO.Administrator = Environ$("USERNAME")
    CSO.AdministrationPassword = EncryptedPassword(NewPassword)
    MsgBox "パスワードを変更しました。", vbInformation, MsgTitle
    Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Sub SignOutAdministrator()
    On Error GoTo ThrowError
    Const ProcedureName As String = "SignOutAdministrator"
    Dim RES As Long
    RES = MsgBox("管理者を離脱しますか？" & vbCrLf & _
        "▲パスワードを消去します。" & vbCrLf & _
        "▲システム再起動します。再起動後にパスワードを再設定してください。", vbYesNo + vbExclamation + vbDefaultButton2, "管理者を離脱")
    If RES <> vbYes Then Exit Sub
    With CSO
        .Administrator = vbNullString
        .AdministrationPassword = vbNullString
        .SkipsPassword = False
    End With
    GateWay.DisconnctFromAddin
Exit Sub
ThrowError:
    CEO.Throw MODULE_NAME, ProcedureName
    Err.Raise Err.Number, , Err.Description
End Sub
Public Function ProtectionPassword() As String
    ProtectionPassword = SHEET_PROTECT_PASSWORD
End Function
Private Function EncryptedPassword(ByVal Password As String) As String
    Dim EncryptKey As String
    EncryptKey = ENCRYPT_KEY
    Do While Len(EncryptKey) < Len(Password)
        EncryptKey = EncryptKey & EncryptKey
    Loop
    Dim Text As String
    Text = Left$(EncryptKey, Len(Password))
    Dim i As Long
    For i = 1 To Len(Password)
        Dim RET As String
        RET = RET & CStr(asc(Mid$(Password, i, 1)) Xor asc(Mid$(Text, i, 1))) & HALF_SPACE
    Next
    EncryptedPassword = RET
End Function
Private Function DecryptedPassword(ByVal EncryptedPassword As String) As String
    Dim Texts() As String
    Texts = Split(EncryptedPassword, HALF_SPACE)
    Dim i As Long
    For i = LBound(Texts) To UBound(Texts) - 1
        Dim RET As String
        RET = RET & Chr$(CLng(Texts(i)) Xor asc(Mid$(ENCRYPT_KEY, i + 1, 1)))
    Next
    DecryptedPassword = RET
End Function
Public Function IsValid() As Boolean
    Const MsgTitle As String = "管理者パスワード照合"
    Dim RES As String
    RES = InputBox("パスワードを入力してください。", MsgTitle)
    If RES Like vbNullString Then Exit Function
    If RES <> DecryptedPassword(CSO.AdministrationPassword) Then
        MsgBox "パスワードが一致しません。", vbExclamation, MsgTitle
        Exit Function
    End If
    IsValid = True
End Function

