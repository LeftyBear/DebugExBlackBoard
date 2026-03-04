VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchoolStructureProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Infrastructure.Config
Option Explicit

Private Type Member
    ConfigList As Scripting.Dictionary
End Type
Private This As Member

Public Sub Initialize(ByVal ConfigList As Scripting.Dictionary)
    Set This.ConfigList = ConfigList
End Sub

Public Function GetLowerGrade() As Long
    GetLowerGrade = GetNumeric("LowerGrade")
End Function

Public Function GetUpperGrade() As Long
    GetUpperGrade = GetNumeric("UpperGrade")
End Function

Public Function GetLowerClass() As Long
    GetLowerClass = GetNumeric("LowerGrade")
End Function

Public Function GetUpperClass() As Long
    GetUpperClass = GetNumeric("UpperGrade")
End Function

Private Function GetNumeric(ByVal KeyName As String) As Long
    Dim SectionName As String
    SectionName = "Common.Status"
    If Not This.ConfigList.Exists(SectionName) Then Err.Raise AppErrNotFoundSection, , "Section not found: " & SectionName
    Dim SectionList As Scripting.Dictionary
    Set SectionList = This.ConfigList.Item(SectionName)
    If Not SectionList.Exists(KeyName) Then Err.Raise AppErrNotFoundKey, , "Key not found:" & KeyName
    GetNumeric = CLng(SectionList.Item(KeyName))
End Function
