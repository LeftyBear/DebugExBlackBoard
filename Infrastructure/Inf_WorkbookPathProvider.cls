VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_WorkbookPathProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.QueryService")
Option Explicit

Public Function GetBaseFolderPath() As String
    Dim CurrentPath As String
    CurrentPath = ThisWorkbook.Path
    Dim ParentPath As String
    ParentPath = GetParentPath(CurrentPath)
    Dim GrandParentPath As String
    GrandParentPath = GetParentPath(ParentPath)
    GetBaseFolderPath = GrandParentPath & charBackSlash
End Function

Private Function GetParentPath(ByVal TargetPath As String) As String
    GetParentPath = VBA.Left$(TargetPath, VBA.InStrRev(TargetPath, charBackSlash) - 1)
End Function
