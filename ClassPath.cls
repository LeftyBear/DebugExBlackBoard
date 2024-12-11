VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SystemObject")
Option Explicit
Public Function ParentPath() As ClassSystemObject
    Set ParentPath = New ClassSystemObject
End Function
Public Function data() As ClassPathData
    Set data = New ClassPathData
End Function
Public Function ico() As ClassPathIco
    Set ico = New ClassPathIco
End Function
Public Function templates() As ClassPathTemplates
    Set templates = New ClassPathTemplates
End Function
Public Function FolderTop() As String
    Dim Sheet As Worksheet
    Set Sheet = ParentPath.BootSheet
    If Sheet Is Nothing Then Exit Function
    FolderTop = Sheet.Range(ParentPath.AddressOfTopFolderPath).Value
End Function
Public Function FolderRoot() As String
    FolderRoot = FSO.BuildPath(Me.FolderTop, "root") & AS_FOLDER
End Function
Public Function FolderData() As String
    FolderData = FSO.BuildPath(Me.FolderRoot, "data") & AS_FOLDER
End Function
Public Function FolderIco() As String
    FolderIco = FSO.BuildPath(Me.FolderRoot, "ico") & AS_FOLDER
End Function
Public Function FolderSrc() As String
    FolderSrc = FSO.BuildPath(Me.FolderRoot, "src") & AS_FOLDER
End Function
Public Function FolderTemplates() As String
    FolderTemplates = FSO.BuildPath(Me.FolderRoot, "templates") & AS_FOLDER
End Function

