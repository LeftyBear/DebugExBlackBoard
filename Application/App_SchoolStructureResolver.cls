VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_SchoolStructureResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Application.Resolver
Option Explicit

Private Type Member
    INILoader As Inf_INILoader
End Type

Private This As Member

Public Sub Initialize(ByVal Loader As Inf_INILoader)
    Set This.INILoader = Loader
End Sub

Public Function Resolve(ByVal SchoolYear As Long) As Dom_SchoolStructure
    If Not (0 < SchoolYear) Then Err.Raise errNotPositiveNumber, , "SchoolYear must be positive"
    Dim AllData As Scripting.Dictionary
    Set AllData = This.INILoader.ReadAll
    Dim SectionName As String
    SectionName = CStr(SchoolYear)
    If Not AllData.Exists(SectionName) Then Err.Raise errNotFoundSection, , "Section not found"
    Dim SectionList As Scripting.Dictionary
    Set SectionList = AllData.Item(SectionName)
    If Not SectionList.Exists("MainStreamGradeCount") Then Err.Raise errNotDefinedStructure, , "MainStreamGradeCount not defined"
    If Not SectionList.Exists("MainStreamClassCount") Then Err.Raise errNotDefinedStructure, , "MainStreamClassCount not defined"
    If Not SectionList.Exists("SpecialGradeCount") Then Err.Raise errNotDefinedStructure, , "SpecialGradeCount not defined"
    If Not SectionList.Exists("SpecialClassCount") Then Err.Raise errNotDefinedStructure, , "SpecialClassCount not defined"
    If Not SectionList.Exists("SubjectCount") Then Err.Raise errNotDefinedStructure, , "SubjectCount not defined"
    If Not SectionList.Exists("PeriodCount") Then Err.Raise errNotDefinedStructure, , "PeriodCount not defined"
    Dim MainStreamGradeCount As Long
    MainStreamGradeCount = CLng(SectionList.Item("MainStreamGradeCount"))
    Dim MainStreamClassCount As Long
    MainStreamClassCount = CLng(SectionList.Item("MainStreamClassCount"))
    Dim SpecialGradeCount As Long
    SpecialGradeCount = CLng(SectionList.Item("SpecialGradeCount"))
    Dim SpecialClassCount As Long
    SpecialClassCount = CLng(SectionList.Item("SpecialClassCount"))
    Dim SubjectCount As Long
    SubjectCount = CLng(SectionList.Item("SubjectCount"))
    Dim PeriodCount As Long
    PeriodCount = CLng(SectionList.Item("PeriodCount"))
    Dim Structure As Dom_SchoolStructure
    Set Structure = New Dom_SchoolStructure
    Structure.Initialize MainStreamGradeCount, MainStreamClassCount, SpecialGradeCount, SpecialClassCount, SubjectCount, PeriodCount
    Set Resolve = Structure
End Function
