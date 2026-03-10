VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ViewEnrollmentFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Formatter")
Option Explicit

Private Enum C
    GradeColumn = 1
    ClassColumn
    MaleColumn
    FemaleColumn
    TotalColumn
End Enum

Public Function Format(ByVal Record As Dom_EnrollmentRecord, ByVal SchoolConfig As Dom_SchoolConfig) As Variant()
    Dim MainStream() As Long
    MainStream = ToTable(Record, Main, SchoolConfig.MainStreamGradeCount, SchoolConfig.MainStreamClassCount)
    Dim SpecialStream() As Long
    SpecialStream = ToTable(Record, Special, SchoolConfig.SpecialGradeCount, SchoolConfig.SpecialClassCount)
    Dim OutputTable() As Variant
    ReDim OutputTable(1 To UBound(MainStream) + UBound(SpecialStream) + 2, 1 To C.TotalColumn)
    OutputTable(1, C.GradeColumn) = "年"
    OutputTable(1, C.ClassColumn) = "組"
    OutputTable(1, C.MaleColumn) = "男"
    OutputTable(1, C.FemaleColumn) = "女"
    OutputTable(1, C.TotalColumn) = "合計"
    Dim R As Long
    R = 2
    Dim Grade As Long
    For Grade = 1 To SchoolConfig.MainStreamGradeCount
        Dim ClassNo As Long
        For ClassNo = 1 To SchoolConfig.MainStreamClassCount
            OutputTable(R, C.GradeColumn) = Grade
            OutputTable(R, C.ClassColumn) = ClassNo
            OutputTable(R, C.MaleColumn) = MainStream(R - 1, C.MaleColumn)
            OutputTable(R, C.FemaleColumn) = MainStream(R - 1, C.FemaleColumn)
            OutputTable(R, C.TotalColumn) = MainStream(R - 1, C.TotalColumn)
            R = R + 1
        Next
        OutputTable(R, C.GradeColumn) = Grade & "年"
        OutputTable(R, C.ClassColumn) = "合計"
        OutputTable(R, C.MaleColumn) = MainStream(R - 1, C.MaleColumn)
        OutputTable(R, C.FemaleColumn) = MainStream(R - 1, C.FemaleColumn)
        OutputTable(R, C.TotalColumn) = MainStream(R - 1, C.TotalColumn)
        Dim GrandMaleTotal As Long
        GrandMaleTotal = GrandMaleTotal + MainStream(R - 1, C.MaleColumn)
        Dim GrandFemaleTotal As Long
        GrandFemaleTotal = GrandFemaleTotal + MainStream(R - 1, C.FemaleColumn)
        R = R + 1
    Next
    For Grade = 1 To SchoolConfig.SpecialGradeCount
        For ClassNo = 1 To SchoolConfig.SpecialClassCount
            OutputTable(R, C.GradeColumn) = "特支"
            OutputTable(R, C.ClassColumn) = ClassNo
            OutputTable(R, C.MaleColumn) = SpecialStream(R - UBound(MainStream) - 1, C.MaleColumn)
            OutputTable(R, C.FemaleColumn) = SpecialStream(R - UBound(MainStream) - 1, C.FemaleColumn)
            OutputTable(R, C.TotalColumn) = SpecialStream(R - UBound(MainStream) - 1, C.TotalColumn)
            R = R + 1
        Next
        OutputTable(R, C.GradeColumn) = "特支"
        OutputTable(R, C.ClassColumn) = "合計"
        OutputTable(R, C.MaleColumn) = SpecialStream(R - UBound(MainStream) - 1, C.MaleColumn)
        OutputTable(R, C.FemaleColumn) = SpecialStream(R - UBound(MainStream) - 1, C.FemaleColumn)
        OutputTable(R, C.TotalColumn) = SpecialStream(R - UBound(MainStream) - 1, C.TotalColumn)
        GrandMaleTotal = GrandMaleTotal + SpecialStream(R - UBound(MainStream) - 1, C.MaleColumn)
        GrandFemaleTotal = GrandFemaleTotal + SpecialStream(R - UBound(MainStream) - 1, C.FemaleColumn)
        R = R + 1
    Next
    OutputTable(R + 1, C.GradeColumn) = "全体"
    OutputTable(R + 1, C.ClassColumn) = "合計"
    OutputTable(R + 1, C.MaleColumn) = GrandMaleTotal
    OutputTable(R + 1, C.FemaleColumn) = GrandFemaleTotal
    OutputTable(R + 1, C.TotalColumn) = GrandMaleTotal + GrandFemaleTotal
    Format = OutputTable
End Function

Private Function ToTable(ByVal Record As Dom_EnrollmentRecord, ByVal StreamType As Dom_StreamType, ByVal GradeCount As Long, ByVal ClassCount As Long) As Long()
    Dim Table() As Long, MaleTotal() As Long, FemaleTotal() As Long, ClassTotal() As Long
    ReDim Table(1 To GradeCount * ClassCount + GradeCount, 1 To C.TotalColumn)
    ReDim MaleTotal(1 To GradeCount)
    ReDim FemaleTotal(1 To GradeCount)
    ReDim ClassTotal(1 To GradeCount, 1 To ClassCount)
    Dim R As Long
    R = 1
    Dim Grade As Long
    For Grade = 1 To GradeCount
        If StreamType = Special Then Grade = 0
        Dim ClassNo As Long
        For ClassNo = 1 To ClassCount
            Dim Gender As Long
            For Gender = Dom_Gender.Male To Dom_Gender.Female
                Dim Value As Long
                Value = Record.Value(StreamType, Grade, ClassNo, Gender)
                If Gender = Dom_Gender.Male Then
                    MaleTotal(Grade) = MaleTotal(Grade) + Value
                    Table(R, C.MaleColumn) = Value
                Else
                    FemaleTotal(Grade) = FemaleTotal(Grade) + Value
                    Table(R, C.FemaleColumn) = Value
                End If
                ClassTotal(Grade, ClassNo) = ClassTotal(Grade, ClassNo) + Value
            Next
            Table(R, C.TotalColumn) = ClassTotal(Grade, ClassNo)
            R = R + 1
        Next
        Table(R, C.MaleColumn) = MaleTotal(Grade)
        Table(R, C.FemaleColumn) = FemaleTotal(Grade)
        Table(R, C.TotalColumn) = MaleTotal(Grade) + FemaleTotal(Grade)
        R = R + 1
    Next
    ToTable = Table
End Function
