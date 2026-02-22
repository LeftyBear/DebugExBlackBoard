VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DirectoryService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Infrastructure.IO
Option Explicit

Public Sub EnsureExists(ByVal FolderPath As String)
    If VBA.Dir(FolderPath, vbDirectory) = vbNullString Then
        VBA.MkDir FolderPath
    End If
End Sub
