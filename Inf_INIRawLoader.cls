VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_INIRawLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Loader"
Option Explicit

Private Type Member
    INIFile As Inf_INIFile
End Type
Private This As Member

Public Sub Initialize(ByVal FilePath As String)
    Set This.INIFile = New Inf_INIFile
    This.INIFile.Load FilePath
End Sub

Public Function ReadAll() As Scripting.Dictionary
    Set ReadAll = This.INIFile.ReadAll
End Function

