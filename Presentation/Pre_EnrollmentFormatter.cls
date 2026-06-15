VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_EnrollmentFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Service.Formatter")
Option Explicit

Private Enum ColumnIndex
    Grade = 1
    ClassNo
    Male
    Female
    Total
    EndColumn
End Enum

Private Const COLUMN_COUNT As Long = ColumnIndex.EndColumn - 1

Public Function Format(ByVal ReadModels As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal MainGradeCount As Long, ByVal MainClassNoCount As Long, ByVal SpecialClassNoCount As Long) As Variant()
    Dim Header() As Variant
    Header = BuildHeader
    Dim Body() As Variant
    Body = BuildBody(ReadModels, Totals, MainGradeCount, MainClassNoCount, SpecialClassNoCount)
    Format = Pre_ArrayUtility.MergeArray(Header, Body)
End Function

Private Function BuildHeader() As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To COLUMN_COUNT)
    Result(1, ColumnIndex.Grade) = "年"
    Result(1, ColumnIndex.ClassNo) = "組"
    Result(1, ColumnIndex.Male) = "男"
    Result(1, ColumnIndex.Female) = "女"
    Result(1, ColumnIndex.Total) = "合計"
    BuildHeader = Result
End Function

Private Function BuildBody(ByVal ReadModels As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal MainGradeCount As Long, ByVal MainClassNoCount As Long, ByVal SpecialClassNoCount As Long) As Variant()
    Dim MainBody() As Variant
    MainBody = BuildBodyPart(ReadModels, Totals, "Main", MainGradeCount, MainClassNoCount)
    Dim SpecialBody() As Variant
    SpecialBody = BuildBodyPart(ReadModels, Totals, "Special", 1, SpecialClassNoCount)
    BuildBody = Pre_ArrayUtility.MergeArray(MainBody, SpecialBody)
End Function

Private Function BuildBodyPart(ByVal ReadModels As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal TypeKey As String, ByVal GradeCount As Long, ByVal ClassNoCount As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To GradeCount * ClassNoCount + GradeCount, 1 To COLUMN_COUNT)
    Dim R As Long
    R = 1
    Dim Grade As Long
    For Grade = 1 To GradeCount
        Dim ClassNo As Long
        For ClassNo = 1 To ClassNoCount
            Result(R, ColumnIndex.Grade) = Grade
            Result(R, ColumnIndex.ClassNo) = ClassNo
            Result(R, ColumnIndex.Male) = GetValue(ReadModels, TypeKey, "Male", Grade, ClassNo)
            Result(R, ColumnIndex.Female) = GetValue(ReadModels, TypeKey, "Female", Grade, ClassNo)
            Result(R, ColumnIndex.Total) = Totals.ClassTotalItem(TypeKey, Grade, ClassNo)
            Dim ClassTotal As Long
            ClassTotal = ClassTotal + Result(R, ColumnIndex.Total)
            R = R + 1
            If (R Mod ClassNoCount + 1) = 0 Then
                Result(R, ColumnIndex.Grade) = Grade & "年"
                Result(R, ColumnIndex.ClassNo) = "計"
                Result(R, ColumnIndex.Male) = Totals.GenderTotalItem(TypeKey, "Male", Grade)
                Result(R, ColumnIndex.Female) = Totals.GenderTotalItem(TypeKey, "Female", Grade)
                Result(R, ColumnIndex.Total) = ClassTotal
                ClassTotal = 0
                R = R + 1
            End If
        Next
    Next
    BuildBodyPart = Result
End Function

Private Function GetValue(ByVal ReadModels As App_EnrollmentReadModels, ByVal TypeKey As String, ByVal Gender As String, ByVal Grade As Long, ByVal ClassNo As Long) As Long
    Dim Model As App_EnrollmentReadModel
    For Each Model In ReadModels
        If IsTarget(Model, TypeKey, Gender, Grade, ClassNo) Then
            GetValue = Model.Value
        End If
    Next
End Function

Private Function IsTarget(ByVal Model As App_EnrollmentReadModel, ByVal TypeKey As String, ByVal Gender As String, ByVal Grade As Long, ByVal ClassNo As Long) As Boolean
    If Model.TypeKey <> TypeKey Then Exit Function
    If Model.Gender <> Gender Then Exit Function
    If Model.Grade <> Grade Then Exit Function
    If Model.ClassNo <> ClassNo Then Exit Function
    IsTarget = True
End Function
