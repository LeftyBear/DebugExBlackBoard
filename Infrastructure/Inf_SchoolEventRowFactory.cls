VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchoolEventRowFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Header As String, ByVal RawText As String) As Inf_SchoolEventRow
    Dim Result As Inf_SchoolEventRow
    Set Result = New Inf_SchoolEventRow
    If Header = jpnIbentoMei Then
        Result.Name = RawText
    End If
    Set Create = Result
End Function
