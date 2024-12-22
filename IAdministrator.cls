VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder("Administrator")
Option Explicit
Public Sub SetPassword()
End Sub
Public Sub SignOut()
End Sub
Public Function ProtectionPassword() As String
End Function
Public Function EncryptedPassword(ByVal Password As String) As String
End Function
Public Function DecryptedPassword(ByVal EncryptedPassword As String) As String
End Function
Public Function IsValid() As Boolean
End Function

