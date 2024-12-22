VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Authorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Authorization")
Option Explicit
Implements IAuthorization
Private Property Let IAuthorization_Administrator(ByVal FilePath As String, ByVal AdministratorName As String)
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        .WriteLine AdministratorName
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Private Property Get IAuthorization_Administrator(ByVal FilePath As String) As String
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForReading)
        IAuthorization_Administrator = .ReadLine
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Private Property Let IAuthorization_AdministrationPassword(ByVal FilePath As String, ByVal Password As String)
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        .WriteLine Password
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Private Property Get IAuthorization_AdministrationPassword(ByVal FilePath As String) As String
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForReading)
        IAuthorization_AdministrationPassword = .ReadLine
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Private Property Let IAuthorization_SkipsAdministrationPassword(ByVal FilePath As String, ByVal State As Boolean)
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForWriting, True)
        .WriteLine CStr(State)
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Private Property Get IAuthorization_SkipsAdministrationPassword(ByVal FilePath As String) As Boolean
    On Error GoTo ThrowError
    With FSO.OpenTextFile(FilePath, ForReading)
        IAuthorization_SkipsAdministrationPassword = CBool(.ReadLine)
ThrowError:
        .Close
        DoEvents
    End With
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Property
Private Function IAuthorization_HasAdministrationPassword(ByVal FilePath As String) As Boolean
    If IAuthorization_AdministrationPassword(FilePath) Like vbNullString Then
        MsgBox "管理者パスワードが設定されていません。" & vbCrLf & _
            "パスワードを設定してください。", vbExclamation, "管理者パスワード設定"
        Exit Function
    End If
    IAuthorization_HasAdministrationPassword = True
End Function
Private Function IAuthorization_IsAdministrator(ByVal FilePath As String) As Boolean
    If Not IAuthorization_Administrator(FilePath) Like Environ$("USERNAME") Then Exit Function
    IAuthorization_IsAdministrator = True
End Function

