VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder("Authorization")
Option Explicit
Public Property Let Administrator(ByVal FilePath As String, ByVal AdministratorName As String)
End Property
Public Property Get Administrator(ByVal FilePath As String) As String
End Property
Public Property Let AdministrationPassword(ByVal FilePath As String, ByVal Password As String)
End Property
Public Property Get AdministrationPassword(ByVal FilePath As String) As String
End Property
Public Property Let SkipsAdministrationPassword(ByVal FilePath As String, ByVal State As Boolean)
End Property
Public Property Get SkipsAdministrationPassword(ByVal FilePath As String) As Boolean
End Property
Public Function IsAdministrator(ByVal FilePath As String) As Boolean
End Function
Public Function HasAdministrationPassword(ByVal FilePath As String) As Boolean
End Function

