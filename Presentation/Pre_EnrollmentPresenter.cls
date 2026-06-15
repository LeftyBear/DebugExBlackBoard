VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pre_EnrollmentPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.Presenter")
Option Explicit
Implements App_IDailyEnrollmentPresenter
Private Enum ColumnIndex
    Grade = 1
    ClassNo
    Male
    Female
    Total
    EndColumn
End Enum
Private Const COLUMN_COUNT As Long = ColumnIndex.EndColumn - 1

Private Function App_IDailyEnrollmentPresenter_CreateViewModel(ByVal ReadModels As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal DTO As App_EnrollmentFormatterDTO) As Pre_EnrollmentViewModel
    Dim Result As Pre_EnrollmentViewModel
    Set Result = New Pre_EnrollmentViewModel
    Result.Table = Format(ReadModels, Totals, DTO)
    Set App_IDailyEnrollmentPresenter_CreateViewModel = Result
End Function

Private Function Format(ByVal ReadModels As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal DTO As App_EnrollmentFormatterDTO) As Variant()
    Dim Header() As Variant
    Header = BuildHeader
    Dim MainBody() As Variant
    MainBody = BuildBody(ReadModels, Totals, DTO, DTO.MainGradeCount, DTO.MainClassNoCount)
    Dim SpecialBody() As Variant
    SpecialBody = BuildBody(ReadModels, Totals, DTO, 1, DTO.SpecialClassNoCount)
    Dim Body() As Variant
    Body = Pre_ArrayUtility.MergeArray(MainBody, SpecialBody)
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

Private Function BuildBody(ByVal ReadModels As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal DTO As App_EnrollmentFormatterDTO, ByVal GradeCount As Long, ByVal ClassNoCount As Long) As Variant()
    Dim RowCount As Long
    RowCount = GradeCount * ClassNoCount + GradeCount
    Dim Result() As Variant
    ReDim Result(1 To RowCount, 1 To COLUMN_COUNT)
    Dim R As Long
    R = 1
    Dim Model As App_EnrollmentReadModel
    For Each Model In ReadModels
        Result(R, ColumnIndex.Grade) = Model.Grade
        If Model.TypeKey = DTO.SpecialKey Then
            Result(R, ColumnIndex.ClassNo) = Model.ClassName
        Else
            Result(R, ColumnIndex.ClassNo) = Model.ClassNo
        End If
        If Model.Gender = DTO.MaleKey Then
            Result(R, ColumnIndex.Male) = Model.Value
        Else
            Result(R, ColumnIndex.Female) = Model.Value
        End If
        Result(R, ColumnIndex.Total) = Totals.ClassTotalItem(Model.TypeKey, Model.Grade, Model.ClassNo)
        Dim ClassTotal As Long
        ClassTotal = ClassTotal + Result(R, ColumnIndex.Total)
        R = R + 1
        If (R Mod ClassNoCount + 1) = 0 Then
            Result(R, ColumnIndex.Grade) = Model.Grade & "年"
            Result(R, ColumnIndex.ClassNo) = "計"
            If Model.Gender = DTO.MaleKey Then
                Result(R, ColumnIndex.Male) = Totals.GenderTotalItem(Model.TypeKey, Model.Gender, Model.Grade)
            Else
                Result(R, ColumnIndex.Female) = Totals.GenderTotalItem(Model.TypeKey, Model.Gender, Model.Grade)
            End If
            Result(R, ColumnIndex.Total) = ClassTotal
            ClassTotal = 0
            R = R + 1
        End If
    Next
    BuildBody = Result
End Function
