VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_PeriodFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Service.Formatter")
Option Explicit

Public Function Format(ByVal Periods As App_Periods, ByVal GradeCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader(GradeCount)
    Dim Body() As Variant
    Body = BuildBody(Periods, GradeCount)
    Format = Pre_ArrayUtility.MergeArray(Header, Body)
End Function

Private Function BuildHeader(ByVal GradeCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To GradeCount)
    Dim C As Long
    For C = 1 To GradeCount
        Result(1, C) = C & "”N"
    Next
    BuildHeader = Result
End Function

Private Function BuildBody(ByVal Periods As App_Periods, ByVal GradeCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To GradeCount)
    If Periods Is Nothing Then BuildBody = Result: Exit Function
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Result(1, Grade) = Periods.Item(Grade)
    Next
    BuildBody = Result
End Function
