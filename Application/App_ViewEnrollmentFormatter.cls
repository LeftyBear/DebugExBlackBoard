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

Public Function Format(ByVal Enrollment As Dom_EnrollmentColumn, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim MainStream() As Long
    MainStream = ToTable(Main, Enrollment, Structure.MainStreamGradeCount, Structure.MainStreamClassCount)
    Dim SpecialStream() As Long
    SpecialStream = ToTable(Special, Enrollment, Structure.SpecialStreamGradeCount, Structure.SpecialStreamClassCount)
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
    For Grade = 1 To Structure.MainStreamGradeCount
        Dim ClassNo As Long
        For ClassNo = 1 To Structure.MainStreamClassCount
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
    For Grade = 1 To Structure.SpecialStreamGradeCount
        For ClassNo = 1 To Structure.SpecialStreamClassCount
            OutputTable(R, C.GradeColumn) = jpnTokushi
            OutputTable(R, C.ClassColumn) = Structure.SpecialStreamClassNameOf(ClassNo)
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

Private Function ToTable(ByVal StreamType As Dom_StreamTypeEnum, ByVal Enrollment As Dom_EnrollmentColumn, ByVal GradeCount As Long, ByVal ClassCount As Long) As Long()
    '@Ignore MultipleDeclarations
    Dim Table() As Long, MaleTotal() As Long, FemaleTotal() As Long, ClassTotal() As Long
    ReDim Table(1 To GradeCount * ClassCount + GradeCount, 1 To C.TotalColumn)
    ReDim MaleTotal(0 To GradeCount)
    ReDim FemaleTotal(0 To GradeCount)
    ReDim ClassTotal(0 To GradeCount, 1 To ClassCount)
    Dim R As Long
    R = 1
    Dim G As Long
    For G = 1 To GradeCount
        Dim Grade As Dom_Grade
        Set Grade = CreateGrade(G)
        Dim No As Long
        For No = 1 To ClassCount
            Dim ClassNo As Dom_ClassNo
            Set ClassNo = CreateClassNo(No)
            Dim Gender As Long
            For Gender = Dom_GenderEnum.Male To Dom_GenderEnum.Female
                Dim Filter As Dom_EnrollmentFilter
                Set Filter = CreateFilter(StreamType, Grade, ClassNo, Gender)
                Dim Value As Long
                Value = Enrollment.GetTotal(Filter)
                If Gender = Dom_GenderEnum.Male Then
                    MaleTotal(G) = MaleTotal(G) + Value
                    Table(R, C.MaleColumn) = Value
                Else
                    FemaleTotal(G) = FemaleTotal(G) + Value
                    Table(R, C.FemaleColumn) = Value
                End If
                ClassTotal(G, No) = ClassTotal(G, No) + Value
            Next
            Table(R, C.TotalColumn) = ClassTotal(G, No)
            R = R + 1
        Next
        Table(R, C.MaleColumn) = MaleTotal(G)
        Table(R, C.FemaleColumn) = FemaleTotal(G)
        Table(R, C.TotalColumn) = MaleTotal(G) + FemaleTotal(G)
        R = R + 1
    Next
    ToTable = Table
End Function

Private Function CreateFilter(ByVal StreamTypeNum As Dom_StreamTypeEnum, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal GenderNum As Long) As Dom_EnrollmentFilter
    Dim StreamType As Dom_StreamType
    Set StreamType = CreateStreamType(StreamTypeNum)
    Dim Gender As Dom_Gender
    Set Gender = CreateGender(GenderNum)
    With New Dom_EnrollmentFilter
        Set CreateFilter = .Create(StreamType, Grade, ClassNo, Gender)
    End With
End Function

Private Function CreateStreamType(ByVal Value As Long) As Dom_StreamType
    With New Dom_StreamType
        Set CreateStreamType = .Create(Value)
    End With
End Function

Private Function CreateGrade(ByVal Value As Long) As Dom_Grade
    With New Dom_Grade
        Set CreateGrade = .Create(Value)
    End With
End Function

Private Function CreateClassNo(ByVal Value As Long) As Dom_ClassNo
    With New Dom_ClassNo
        Set CreateClassNo = .Create(Value)
    End With
End Function

Private Function CreateGender(ByVal Value As Long) As Dom_Gender
    With New Dom_Gender
        Set CreateGender = .Create(Value)
    End With
End Function
