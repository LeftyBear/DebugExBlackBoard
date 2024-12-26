VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Administrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Administrator")
Option Explicit
Implements IAdministrator
Private Authorization               As IAuthorization
Private Directory                   As IDirectory
Private Const PROTECTION_PASSWORD   As String = "pass3225"
Private Const ENCRYPT_KEY           As String = "f4jO1cYz"
Private Sub IAdministrator_SignOut()
    Dim RES As Long
    RES = MsgBox("管理者を離脱しますか？" & vbCrLf & _
        "▲パスワードを消去します。" & vbCrLf & _
        "▲システム再起動します。再起動後にパスワードを再設定してください。", vbYesNo + vbExclamation + vbDefaultButton2, "管理者を離脱")
    If RES <> vbYes Then Exit Sub
    Set Authorization = New Authorization
    With Authorization
        .Administrator(System.Directory.AdministorFile) = vbNullString
        .AdministrationPassword(System.Directory.PasswordFile) = vbNullString
        .SkipsAdministrationPassword(System.Directory.PasswordSkipFile) = False
    End With
End Sub
Private Sub IAdministrator_SetPassword()
    Const MsgTitle As String = "管理者パスワード変更"
    Do
        Dim NewPassword As String
        NewPassword = InputBox("新しいパスワードを入力してください。", MsgTitle)
        If NewPassword Like vbNullString Then Exit Sub
        Dim ConfirmPassword As String
        ConfirmPassword = InputBox("確認のためにもう一度パスワードを入力してください。", MsgTitle)
        If ConfirmPassword Like vbNullString Then Exit Sub
        If NewPassword Like ConfirmPassword Then Exit Do
        MsgBox "パスワードが一致しません。はじめからやり直してください。", vbExclamation, MsgTitle
    Loop
    Set Authorization = New Authorization
    With Authorization
        Set Directory = New Directory
        .Administrator(Directory.AdministorFile) = Environ$("USERNAME")
        .AdministrationPassword(Directory.PasswordFile) = IAdministrator_EncryptedPassword(NewPassword)
    End With
    MsgBox "パスワードを変更しました。", vbInformation, MsgTitle
End Sub
Private Function IAdministrator_EncryptedPassword(ByVal Password As String) As String
    Dim RET As String
    Dim EncryptKey As String
    EncryptKey = ENCRYPT_KEY
    Do While Len(EncryptKey) < Len(Password)
        EncryptKey = EncryptKey & EncryptKey
    Loop
    Dim BUF As String
    BUF = Left$(EncryptKey, Len(Password))
    Dim i As Long
    For i = 1 To Len(Password)
        RET = RET & CStr(Asc(Mid$(Password, i, 1)) Xor Asc(Mid$(BUF, i, 1))) & HALF_SPACE
    Next
    IAdministrator_EncryptedPassword = RET
End Function
Private Function IAdministrator_DecryptedPassword(ByVal EncryptedPassword As String) As String
    Dim RET As String
    Dim BUF() As String
    BUF = Split(EncryptedPassword, HALF_SPACE)
    Dim i As Long
    For i = LBound(BUF) To UBound(BUF) - 1
        RET = RET & Chr$(CLng(BUF(i)) Xor Asc(Mid$(ENCRYPT_KEY, i + 1, 1)))
    Next
    IAdministrator_DecryptedPassword = RET
End Function
Private Function IAdministrator_ProtectionPassword() As String
    IAdministrator_ProtectionPassword = PROTECTION_PASSWORD
End Function
Private Function IAdministrator_IsValid() As Boolean
    Const MsgTitle As String = "管理者パスワード照合"
    Dim RES As String
    RES = InputBox("パスワードを入力してください。", MsgTitle)
    If RES Like vbNullString Then Exit Function
    Set Authorization = New Authorization
    Set Directory = New Directory
    If Not RES Like IAdministrator_DecryptedPassword(Authorization.AdministrationPassword(Directory.PasswordFile)) Then
        MsgBox "パスワードが一致しません。", vbExclamation, MsgTitle: Exit Function
    End If
    IAdministrator_IsValid = True
End Function

