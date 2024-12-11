VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPathIco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassPath
    Set ParentPath = New ClassPath
End Function
Public Function FileMousecursor() As String
    FileMousecursor = FSO.BuildPath(Me.ParentPath.FolderIco, "cursor.ico")
End Function

