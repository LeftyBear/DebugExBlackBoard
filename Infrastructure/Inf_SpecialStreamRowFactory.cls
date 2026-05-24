VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SpecialStreamRowFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Private Const SPECIAL_NAME As String = "“ÁŽx–¼"

Public Function Create(ByVal Header As String, ByVal RawText As String) As Inf_SpecialStreamRow
    Dim Result As Inf_SpecialStreamRow
    Set Result = New Inf_SpecialStreamRow
    If Header = SPECIAL_NAME Then
        Result.Name = RawText
    End If
    Set Create = Result
End Function
