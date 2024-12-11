VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPathCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassPathData
    Set ParentPath = New ClassPathData
End Function
Public Function FileAdministrator() As String
    FileAdministrator = FSO.BuildPath(Me.ParentPath.FolderCache, "Administer.txt")
End Function
Public Function FilePassword() As String
    FilePassword = FSO.BuildPath(Me.ParentPath.FolderCache, "Password.txt")
End Function
Public Function FilePasswordSkip() As String
    FilePasswordSkip = FSO.BuildPath(Me.ParentPath.FolderCache, "PasswordSkip.txt")
End Function
Public Function FileOnTime() As String
    FileOnTime = FSO.BuildPath(Me.ParentPath.FolderCache, "OnTime.log")
End Function

