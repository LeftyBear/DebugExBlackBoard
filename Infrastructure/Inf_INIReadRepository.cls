VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_INIReadRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Loader"
Option Explicit

Private Type Member
    Reader As Inf_INIReader
End Type

Private This As Member

Public Sub Initialize(ByVal Reader As Inf_INIReader)
    Set This.Reader = Reader
End Sub

Public Function ReadAll(ByVal FilePath As String) As String()
    ReadAll = This.Reader.ReadAll(FilePath)
End Function
