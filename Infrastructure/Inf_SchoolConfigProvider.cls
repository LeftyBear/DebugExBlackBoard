VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SchoolConfigProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Infrastructure.Config
Option Explicit

Private Type Member
    INILoader As Inf_INILoader
End Type
Private This As Member

Public Sub Initialize(ByVal Loader As Inf_INILoader)
    Set This.INILoader = Loader
End Sub

Public Function GetLowerGrade(ByVal SchoolYear As Long) As Long
    GetLowerGrade = GetValue(SchoolYear, "LowerGrade")
End Function

Public Function GetUpperGrade(ByVal SchoolYear As Long) As Long
    GetUpperGrade = GetValue(SchoolYear, "UpperGrade")
End Function

Public Function GetLowerClass(ByVal SchoolYear As Long) As Long
    GetLowerClass = GetValue(SchoolYear, "LowerGrade")
End Function

Public Function GetUpperClass(ByVal SchoolYear As Long) As Long
    GetUpperClass = GetValue(SchoolYear, "UpperGrade")
End Function

Private Function GetValue(ByVal SchoolYear As Long, ByVal KeyName As String) As Long
    If Not (0 < SchoolYear) Then Err.Raise App_ErrNum.NotPositiveNumber, , "SchoolYear must be positive."
    Dim AllData As Scripting.Dictionary
    Set AllData = This.INILoader.ReadAll
    Dim SectionName As String
    SectionName = CStr(SchoolYear)
    If Not AllData.Exists(SectionName) Then Err.Raise App_ErrNum.NotFoundSection, , "Section not found."
    Dim SectionDictionary As Scripting.Dictionary
    Set SectionDictionary = AllData.Item(SectionName)
    If Not SectionDictionary.Exists(KeyName) Then Err.Raise App_ErrNum.NotFoundKey, , "Key not found."
    GetValue = CLng(SectionDictionary.Item(KeyName))
End Function
