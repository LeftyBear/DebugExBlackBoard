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

Public Function Resolve(ByRef Source() As String) As Dom_LimitValue
    Dim Structure As Dom_LimitValue
    Set Structure = New Dom_LimitValue
'    Structure.Initialize MainStreamGradeCount, MainStreamClassCount, SpecialGradeCount, SpecialClassCount, PeriodCount
    Set Resolve = Structure
End Function

Private Sub AddCommonStatus(ByVal SectionList As Scripting.Dictionary, ByRef MainStreamGradeCount As Long, ByRef MainStreamClassCount As Long, ByRef SpecialGradeCount As Long, ByRef SpecialClassCount As Long)
    If Not SectionList.Exists("MainStreamGradeCount") Then Err.Raise AppErrNotDefinedStructure, , "MainStreamGradeCount not defined"
    If Not SectionList.Exists("MainStreamClassCount") Then Err.Raise AppErrNotDefinedStructure, , "MainStreamClassCount not defined"
    If Not SectionList.Exists("SpecialGradeCount") Then Err.Raise AppErrNotDefinedStructure, , "SpecialGradeCount not defined"
    If Not SectionList.Exists("SpecialClassCount") Then Err.Raise AppErrNotDefinedStructure, , "SpecialClassCount not defined"
    MainStreamGradeCount = CLng(SectionList.Item("MainStreamGradeCount"))
    MainStreamClassCount = CLng(SectionList.Item("MainStreamClassCount"))
    SpecialGradeCount = CLng(SectionList.Item("SpecialGradeCount"))
    SpecialClassCount = CLng(SectionList.Item("SpecialClassCount"))
End Sub

Private Sub AddClassHourStatus(ByVal SectionList As Scripting.Dictionary, ByRef SubjectCount As Long, ByRef PeriodCount As Long)
    If Not SectionList.Exists("PeriodCount") Then Err.Raise AppErrNotDefinedStructure, , "PeriodCount not defined"
    PeriodCount = CLng(SectionList.Item("PeriodCount"))
End Sub
