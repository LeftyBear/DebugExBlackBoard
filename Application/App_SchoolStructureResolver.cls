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

Public Function Resolve(ByVal SchoolYearKey As App_SchoolYearKey) As Dom_SchoolStructure
    Dim SchoolYear As Long
    SchoolYear = SchoolYearKey.Value
    If Not (0 < SchoolYear) Then Err.Raise App_ErrNum.NotPositiveNumber, , "SchoolYear must be positive"
    Dim AllData As Scripting.Dictionary
    Set AllData = This.INILoader.ReadAll
    Dim SectionName As String
    SectionName = CStr(SchoolYear)
    If Not AllData.Exists(SectionName) Then Err.Raise App_ErrNum.NotFoundSection, , "Section not found"
    Dim SectionList As Scripting.Dictionary
    Set SectionList = AllData.Item(SectionName)
    If Not SectionList.Exists("MainStreamGradeCount") Then Err.Raise App_ErrNum.NotDefinedStructure, , "MainStreamGradeCount not defined"
    If Not SectionList.Exists("MainStreamClassCount") Then Err.Raise App_ErrNum.NotDefinedStructure, , "MainStreamClassCount not defined"
    If Not SectionList.Exists("SpecialGradeCount") Then Err.Raise App_ErrNum.NotDefinedStructure, , "SpecialGradeCount not defined"
    If Not SectionList.Exists("SpecialClassCount") Then Err.Raise App_ErrNum.NotDefinedStructure, , "SpecialClassCount not defined"
    If Not SectionList.Exists("SubjectCount") Then Err.Raise App_ErrNum.NotDefinedStructure, , "SubjectCount not defined"
    If Not SectionList.Exists("PeriodCount") Then Err.Raise App_ErrNum.NotDefinedStructure, , "PeriodCount not defined"
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
