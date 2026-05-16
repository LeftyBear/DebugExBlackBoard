VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Factory     As Dom_ValueObjectFactory
    Calculator  As App_EnrollmentSummaryCalculator
End Type

Private This As Member

Private Enum C
    GradeColumn = 1
    ClassColumn
    MaleColumn
    FemaleColumn
    TotalColumn
End Enum

Private Sub Class_Initialize()
    Set This.Factory = New Dom_ValueObjectFactory
    Set This.Calculator = New App_EnrollmentSummaryCalculator
End Sub
'
'Public Sub Initialize(ByVal Calculator As App_EnrollmentSummaryCalculator)
'    Set This.Calculator = Calculator
'End Sub

Public Function Format(ByVal Aggregate As Dom_DailyEnrollment, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim MainStreamGradeCount As Long
    MainStreamGradeCount = Structure.UpperValue.Item.MainStreamGradeCount
    Dim MainStreamClassCount As Long
    MainStreamClassCount = Structure.UpperValue.Item.MainStreamClassCount
    Dim SpecialStreamGradeCount As Long
    SpecialStreamGradeCount = 1
    Dim SpecialStreamClassCount As Long
    SpecialStreamClassCount = Structure.SpecialStream.Count
    Dim MainStream() As Long
    MainStream = ToTable(Main, Aggregate, MainStreamGradeCount, MainStreamClassCount)
    Dim SpecialStream() As Long
    SpecialStream = ToTable(Special, Aggregate, SpecialStreamGradeCount, SpecialStreamClassCount)
    Dim OutputTable() As Variant
    ReDim OutputTable(1 To UBound(MainStream) + UBound(SpecialStream) + 2, 1 To C.TotalColumn)
    OutputTable(1, C.GradeColumn) = jpnNen
    OutputTable(1, C.ClassColumn) = jpnKumi
    OutputTable(1, C.MaleColumn) = jpnOtoko
    OutputTable(1, C.FemaleColumn) = jpnOnna
    OutputTable(1, C.TotalColumn) = jpnGoukei
    Dim R As Long
    R = 2
    Dim Grade As Long
    For Grade = 1 To MainStreamGradeCount
        Dim ClassNo As Long
        For ClassNo = 1 To MainStreamClassCount
            OutputTable(R, C.GradeColumn) = Grade
            OutputTable(R, C.ClassColumn) = ClassNo
            OutputTable(R, C.MaleColumn) = MainStream(R - 1, C.MaleColumn)
            OutputTable(R, C.FemaleColumn) = MainStream(R - 1, C.FemaleColumn)
            OutputTable(R, C.TotalColumn) = MainStream(R - 1, C.TotalColumn)
            R = R + 1
        Next
        OutputTable(R, C.GradeColumn) = Grade & jpnNen
        OutputTable(R, C.ClassColumn) = jpnGoukei
        OutputTable(R, C.MaleColumn) = MainStream(R - 1, C.MaleColumn)
        OutputTable(R, C.FemaleColumn) = MainStream(R - 1, C.FemaleColumn)
        OutputTable(R, C.TotalColumn) = MainStream(R - 1, C.TotalColumn)
        Dim GrandMaleTotal As Long
        GrandMaleTotal = GrandMaleTotal + MainStream(R - 1, C.MaleColumn)
        Dim GrandFemaleTotal As Long
        GrandFemaleTotal = GrandFemaleTotal + MainStream(R - 1, C.FemaleColumn)
        R = R + 1
    Next
    For Grade = 1 To SpecialStreamGradeCount
        For ClassNo = 1 To SpecialStreamClassCount
            OutputTable(R, C.GradeColumn) = jpnTokushi
            OutputTable(R, C.ClassColumn) = Structure.SpecialStream.FindByIndex(ClassNo).Value
            OutputTable(R, C.MaleColumn) = SpecialStream(R - UBound(MainStream) - 1, C.MaleColumn)
            OutputTable(R, C.FemaleColumn) = SpecialStream(R - UBound(MainStream) - 1, C.FemaleColumn)
            OutputTable(R, C.TotalColumn) = SpecialStream(R - UBound(MainStream) - 1, C.TotalColumn)
            R = R + 1
        Next
        OutputTable(R, C.GradeColumn) = jpnTokushi
        OutputTable(R, C.ClassColumn) = jpnGoukei
        OutputTable(R, C.MaleColumn) = SpecialStream(R - UBound(MainStream) - 1, C.MaleColumn)
        OutputTable(R, C.FemaleColumn) = SpecialStream(R - UBound(MainStream) - 1, C.FemaleColumn)
        OutputTable(R, C.TotalColumn) = SpecialStream(R - UBound(MainStream) - 1, C.TotalColumn)
        GrandMaleTotal = GrandMaleTotal + SpecialStream(R - UBound(MainStream) - 1, C.MaleColumn)
        GrandFemaleTotal = GrandFemaleTotal + SpecialStream(R - UBound(MainStream) - 1, C.FemaleColumn)
        R = R + 1
    Next
    OutputTable(R, C.GradeColumn) = jpnZentai
    OutputTable(R, C.ClassColumn) = jpnGoukei
    OutputTable(R, C.MaleColumn) = GrandMaleTotal
    OutputTable(R, C.FemaleColumn) = GrandFemaleTotal
    OutputTable(R, C.TotalColumn) = GrandMaleTotal + GrandFemaleTotal
    Format = OutputTable
End Function

Private Function ToTable(ByVal TypeCode As Dom_StreamTypePolicy, ByVal Aggregate As Dom_DailyEnrollment, ByVal GradeCount As Long, ByVal ClassCount As Long) As Long()
    '@Ignore MultipleDeclarations
    Dim Table() As Long, MaleTotal() As Long, FemaleTotal() As Long, ClassTotal() As Long
    ReDim Table(1 To GradeCount * ClassCount + GradeCount, 1 To C.TotalColumn)
    If Aggregate Is Nothing Then GoTo Finalize
    ReDim MaleTotal(0 To GradeCount)
    ReDim FemaleTotal(0 To GradeCount)
    ReDim ClassTotal(0 To GradeCount, 1 To ClassCount)
    Dim StreamType As Dom_StreamType
    Set StreamType = This.Factory.CreateStreamType(TypeCode)
    Dim R As Long
    R = 1
    Dim G As Long
    For G = 1 To GradeCount
        Dim Grade As Dom_Grade
        Set Grade = This.Factory.CreateGrade(G)
        Dim No As Long
        For No = 1 To ClassCount
            Dim ClassNo As Dom_ClassNo
            Set ClassNo = This.Factory.CreateClassNo(No)
            Dim GenderNo As Long
            For GenderNo = Dom_GenderTypePolicy.Male To Dom_GenderTypePolicy.Female
                Dim Gender As Dom_Gender
                Set Gender = This.Factory.CreateGender(GenderNo)
                Dim Condition As App_EnrollmentIDCondition
                Set Condition = This.Factory.CreateEnrollmentFilterCondition(StreamType, Grade, ClassNo, Gender)
                Dim EnrollmentValue As Dom_EnrollmentValue
                Set EnrollmentValue = This.Calculator.SumClassOf(Aggregate, Condition)
                If Gender.Value = Dom_GenderTypePolicy.Male Then
                    MaleTotal(G) = MaleTotal(G) + EnrollmentValue.Value
                    Table(R, C.MaleColumn) = EnrollmentValue.Value
                Else
                    FemaleTotal(G) = FemaleTotal(G) + EnrollmentValue.Value
                    Table(R, C.FemaleColumn) = EnrollmentValue.Value
                End If
                ClassTotal(G, No) = ClassTotal(G, No) + EnrollmentValue.Value
            Next
            Table(R, C.TotalColumn) = ClassTotal(G, No)
            R = R + 1
        Next
        Table(R, C.MaleColumn) = MaleTotal(G)
        Table(R, C.FemaleColumn) = FemaleTotal(G)
        Table(R, C.TotalColumn) = MaleTotal(G) + FemaleTotal(G)
        R = R + 1
    Next
Finalize:
    ToTable = Table
End Function
