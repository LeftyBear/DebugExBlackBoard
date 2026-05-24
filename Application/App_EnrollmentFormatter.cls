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

Private Enum C_
    Grade = 1
    ClassNo
    Male
    Female
    Total
    EndColumn
End Enum

Private Const COLUMN_COUNT As Long = C_.EndColumn - 1

Public Function Format(ByVal Models As App_EnrollmentReadModels, ByVal Structure As Dom_SchoolStructure) As Variant()
    Dim Header() As Variant
    Header = BuildHeader
    Dim MainUpperGrade As Long
    MainUpperGrade = Structure.MainStream.Item.UpperGrade
    Dim MainUpperClassNo As Long
    MainUpperClassNo = Structure.MainStream.Item.UpperClassNo
    Dim MainTable() As Long
    MainTable = BuildGenderTable(Models, "Main", MainUpperGrade, MainUpperClassNo)
    Dim SpecialUpperClassNo As Long
'    SpecialUpperClassNo = Structure.SpecialStream.Count
    Dim SpecialTable() As Long
    SpecialTable = BuildGenderTable(Models, "Special", 1, SpecialUpperClassNo)
    Dim Body() As Variant
    Body = BuildBody(MainTable, SpecialTable, MainUpperClassNo, SpecialUpperClassNo)
    Format = BuildTable(Header, Body)
End Function

Private Sub SumGender(ByRef Base() As Long, ByRef MaleSummary() As Long, ByRef FemaleSummary() As Long)
    Dim R As Long
    For R = 1 To UBound(MaleSummary)
        Base(R, C_.Male) = MaleSummary(R)
        Base(R, C_.Female) = FemaleSummary(R)
        Base(R, C_.Total) = MaleSummary(R) + FemaleSummary(R)
    Next
End Sub

Private Sub AddClassName(ByRef Base() As Long, ByRef Grades() As Long, ByRef ClassNumbers() As Long)
    Dim R As Long
    For R = 1 To UBound(Base)
        Base(R, C_.Grade) = Grades(R)
        Base(R, C_.ClassNo) = ClassNumbers(R)
    Next
End Sub

Private Function BuildHeader() As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To COLUMN_COUNT)
    Result(1, C_.Grade) = jpnNen
    Result(1, C_.ClassNo) = jpnKumi
    Result(1, C_.Male) = jpnOtoko
    Result(1, C_.Female) = jpnOnna
    Result(1, C_.Total) = jpnGoukei
    BuildHeader = Result
End Function

Private Function BuildTable(ByRef Header() As Variant, ByRef Body() As Variant) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To UBound(Header) + UBound(Body), 1 To COLUMN_COUNT)
    Dim R As Long
    For R = 1 To UBound(Header)
        Dim C As Long
        For C = 1 To COLUMN_COUNT
            Result(R, C) = Header(R, C)
        Next
    Next
    For R = UBound(Header) + 1 To UBound(Body)
        For C = 1 To COLUMN_COUNT
            Result(R, C) = Body(R, C)
        Next
    Next
    BuildTable = Result
End Function

Private Function BuildBody(ByRef MainTable() As Long, ByRef SpecialTable() As Long, ByVal MainUpperClassNo As Long, ByVal SpecialUpperClassNo As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To UBound(MainTable) + UBound(SpecialTable), 1 To COLUMN_COUNT)
    Dim R As Long
    For R = 1 To UBound(MainTable)
        Dim C As Long
        For C = 1 To COLUMN_COUNT
            If (R Mod MainUpperClassNo + 1) = 0 Then
                Result(R, C_.Total) = jpnGoukei
            Else
                Result(R, C) = MainTable(R, C)
            End If
        Next
    Next
    For R = UBound(MainTable) + 1 To UBound(SpecialTable)
        For C = 1 To COLUMN_COUNT
            If (R Mod SpecialUpperClassNo + 1) = 0 Then
                Result(R, C_.Total) = jpnGoukei
            Else
                Result(R, C) = SpecialTable(R, C)
            End If
        Next
    Next
    BuildBody = Result
End Function

Private Function BuildGenderTable(ByVal Models As App_EnrollmentReadModels, ByVal TypeCode As String, ByVal UpperGrade As Long, ByVal UpperClassNo As Long) As Long()
    Dim MaleSummary() As Long
    MaleSummary = GetMaleSummary(Models, TypeCode, "Male", UpperGrade, UpperClassNo)
    Dim FemaleSummary() As Long
    FemaleSummary = GetFemaleSummary(Models, TypeCode, "Female", UpperGrade, UpperClassNo)
    Dim Table() As Long
    ReDim Table(1 To UpperGrade * UpperClassNo + UpperGrade, 1 To COLUMN_COUNT)
    SumGender Table, MaleSummary, FemaleSummary
    Dim Grades() As Long
    Grades = GetGrades(UpperGrade, UpperClassNo)
    Dim ClassNumbers() As Long
    ClassNumbers = GetClassNumbers(UpperGrade, UpperClassNo)
    AddClassName Table, Grades, ClassNumbers
    BuildGenderTable = Table
End Function

Private Function GetMaleSummary(ByVal Models As App_EnrollmentReadModels, ByVal TypeCode As String, ByVal Gender As String, ByVal UpperGrade As Long, ByVal UpperClassNo As Long) As Long()
    Dim Values() As Long
    Values = FilterBy(Models, TypeCode, "Male", UpperGrade, UpperClassNo)
    GetMaleSummary = CalculateGenderTotal(Values, UpperGrade, UpperClassNo)
End Function

Private Function GetFemaleSummary(ByVal Models As App_EnrollmentReadModels, ByVal TypeCode As String, ByVal Gender As String, ByVal UpperGrade As Long, ByVal UpperClassNo As Long) As Long()
    Dim Values() As Long
    Values = FilterBy(Models, TypeCode, "Female", UpperGrade, UpperClassNo)
    GetFemaleSummary = CalculateGenderTotal(Values, UpperGrade, UpperClassNo)
End Function

Private Function FilterBy(ByVal Models As App_EnrollmentReadModels, ByVal TypeCode As String, ByVal Gender As String, ByVal UpperGrade As Long, ByVal UpperClassNo As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To UpperGrade, 1 To UpperClassNo)
    Dim Model As App_EnrollmentReadModel
    For Each Model In Models
        If Model.TypeCode = TypeCode Then
            If Model.Gender = Gender Then
                Result(Model.Grade, Model.ClassNo) = Model.Value
            End If
        End If
    Next
    FilterBy = Result
End Function

Private Function CalculateGenderTotal(ByRef Values() As Long, ByVal UpperGrade As Long, ByVal UpperClassNo As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To UpperGrade * UpperClassNo + UpperGrade)
    Dim Grade As Long
    For Grade = 1 To UpperGrade
        Dim ClassNo As Long
        For ClassNo = 1 To UpperClassNo
            Dim R As Long
            If R < UpperClassNo + 1 Then
                Result(R) = Values(Grade, ClassNo)
                Dim Total As Long
                Total = Total + Values(Grade, ClassNo)
            End If
            R = R + 1
            If (R Mod UpperClassNo + 1) = 0 Then
                Result(R) = Total
                Total = 0
                Dim GrandTotal As Long
                GrandTotal = GrandTotal + Total
            End If
        Next
    Next
    Result(R) = GrandTotal
    CalculateGenderTotal = Result
End Function

Private Function GetGrades(ByVal UpperGrade As Long, ByVal UpperClassNo As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To UpperGrade * UpperClassNo + UpperGrade)
    Dim Grade As Long
    For Grade = 1 To UpperGrade
        Dim ClassNo As Long
        For ClassNo = 1 To UpperClassNo
            Dim R As Long
            If R < UpperClassNo + 1 Then
                Result(R) = Grade
            End If
            R = R + 1
        Next
    Next
    GetGrades = Result
End Function

Private Function GetClassNumbers(ByVal UpperGrade As Long, ByVal UpperClassNo As Long) As Long()
    Dim Result() As Long
    ReDim Result(1 To UpperGrade * UpperClassNo + UpperGrade)
    Dim Grade As Long
    For Grade = 1 To UpperGrade
        Dim ClassNo As Long
        For ClassNo = 1 To UpperClassNo
            Dim R As Long
            If R < UpperClassNo + 1 Then
                Result(R) = ClassNo
            End If
            R = R + 1
        Next
    Next
    GetClassNumbers = Result
End Function
