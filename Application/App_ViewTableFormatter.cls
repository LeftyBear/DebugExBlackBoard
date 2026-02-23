VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ViewTableFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Formatter")
Option Explicit

Public Sub FormatCombined(ByVal Summary As Dom_EnrollmentSummary, ByVal SchoolStructure As Dom_SchoolStructure, ByRef OutputArray() As Variant)
    Dim MainArray() As Variant
    Dim SpecialArray() As Variant
    Dim MainRows As Long
    Dim SpecialRows As Long
    Dim TotalRows As Long
    Dim RowIndex As Long
    Dim ColIndex As Long
    Format Summary, SchoolStructure, Dom_StreamType.Main, MainArray
    Format Summary, SchoolStructure, Dom_StreamType.Special, SpecialArray
    MainRows = UBound(MainArray, 1)
    SpecialRows = UBound(SpecialArray, 1)
    TotalRows = MainRows + 1 + SpecialRows
    ReDim OutputArray(1 To TotalRows, 1 To 5)
    For RowIndex = 1 To MainRows
        For ColIndex = 1 To 5
            OutputArray(RowIndex, ColIndex) = MainArray(RowIndex, ColIndex)
        Next
    Next
    For ColIndex = 1 To 5
        OutputArray(MainRows + 1, ColIndex) = vbNullString
    Next
    For RowIndex = 1 To SpecialRows
        For ColIndex = 1 To 5
            OutputArray(MainRows + 1 + RowIndex, ColIndex) = SpecialArray(RowIndex, ColIndex)
        Next
    Next
End Sub

Private Sub Format(ByVal Summary As Dom_EnrollmentSummary, ByVal SchoolStructure As Dom_SchoolStructure, ByVal StreamType As Dom_StreamType, ByRef OutputArray() As Variant)
    Dim Grade As Long
    Dim ClassNo As Long
    Dim GradeCount As Long
    Dim ClassCount As Long
    Dim RowCount As Long
    Dim RowIndex As Long
    Dim Male As Long
    Dim Female As Long
    Dim GradeMaleTotal As Long
    Dim GradeFemaleTotal As Long
    Dim GrandMaleTotal As Long
    Dim GrandFemaleTotal As Long
    Select Case StreamType
    Case Main
        GradeCount = SchoolStructure.MainStreamGradeCount
        ClassCount = SchoolStructure.MainStreamClassCount
    Case Special
        GradeCount = SchoolStructure.SpecialGradeCount
        ClassCount = SchoolStructure.SpecialClassCount
    End Select
    RowCount = 1
    For Grade = 1 To GradeCount
        RowCount = RowCount + ClassCount + 1
    Next
    RowCount = RowCount + 1
    ReDim OutputArray(1 To RowCount, 1 To 5)
    OutputArray(1, 1) = "Grade"
    OutputArray(1, 2) = "Class"
    OutputArray(1, 3) = "Male"
    OutputArray(1, 4) = "Female"
    OutputArray(1, 5) = "Total"
    RowIndex = 1
    For Grade = 1 To GradeCount
        GradeMaleTotal = 0
        GradeFemaleTotal = 0
        For ClassNo = 1 To ClassCount
            Male = Summary.GetTotal(StreamType, Grade, ClassNo, Dom_Gender.Male)
            Female = Summary.GetTotal(StreamType, Grade, ClassNo, Dom_Gender.Female)
            RowIndex = RowIndex + 1
            OutputArray(RowIndex, 1) = Grade
            OutputArray(RowIndex, 2) = ClassNo
            OutputArray(RowIndex, 3) = Male
            OutputArray(RowIndex, 4) = Female
            OutputArray(RowIndex, 5) = Male + Female
            GradeMaleTotal = GradeMaleTotal + Male
            GradeFemaleTotal = GradeFemaleTotal + Female
            GrandMaleTotal = GrandMaleTotal + Male
            GrandFemaleTotal = GrandFemaleTotal + Female
        Next
        RowIndex = RowIndex + 1
        OutputArray(RowIndex, 1) = Grade
        OutputArray(RowIndex, 2) = "çáåv"
        OutputArray(RowIndex, 3) = GradeMaleTotal
        OutputArray(RowIndex, 4) = GradeFemaleTotal
        OutputArray(RowIndex, 5) = GradeMaleTotal + GradeFemaleTotal
    Next
    RowIndex = RowIndex + 1
    OutputArray(RowIndex, 1) = "ëSëÃ"
    OutputArray(RowIndex, 2) = "çáåv"
    OutputArray(RowIndex, 3) = GrandMaleTotal
    OutputArray(RowIndex, 4) = GrandFemaleTotal
    OutputArray(RowIndex, 5) = GrandMaleTotal + GrandFemaleTotal
End Sub
