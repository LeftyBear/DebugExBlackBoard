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
Private Enum INDEX
    GRADE_COLUMN = 1
    CLASSNO_COLUMN
    MALE_COLUMN
    FEMALE_COLUMN
    TOTAL_COLUMN
    ADDITIONAL_COLUMN
End Enum
Private Type Member
    Grade   As Long
    ClassNo As Long
    Male    As Long
    Female  As Long
    Total   As Long
    Count   As Long
End Type

Private C As Member

Private Sub Class_Initialize()
    C.Grade = INDEX.GRADE_COLUMN
    C.ClassNo = INDEX.CLASSNO_COLUMN
    C.Male = INDEX.MALE_COLUMN
    C.Female = INDEX.FEMALE_COLUMN
    C.Total = INDEX.TOTAL_COLUMN
    C.Count = INDEX.ADDITIONAL_COLUMN - 1
End Sub

Private Function Format(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal MainStreamRM As App_MainStreamReadModel, ByVal SpecialStreamRMs As App_SpecialStreamReadModels) As Variant()
    Dim UpperColumn As Long
    UpperColumn = C.Count
    Dim MainGradeCount As Long
    MainGradeCount = MainStreamRM.UpperGrade
    Dim MainClassNoCount As Long
    MainClassNoCount = MainStreamRM.UpperClassNo
    Dim SpecialClassNoCount As Long
    SpecialClassNoCount = SpecialStreamRMs.Count
    Dim Header() As Variant
    Header = BuildHeader(UpperColumn)
    Dim Body() As Variant
    Body = BuildBody(EnrollmentRMs, Totals, MainGradeCount, MainClassNoCount, SpecialClassNoCount)
    Dim Bottom() As Variant
    Bottom = BuildBottom(Totals, UpperColumn)
    Format = Pre_ArrayUtility.MergeVertically(Header, Body)
End Function

Private Function BuildHeader(ByVal UpperColumn As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To UpperColumn)
    Result(1, C.Grade) = "年"
    Result(1, C.ClassNo) = "組"
    Result(1, C.Male) = "男"
    Result(1, C.Female) = "女"
    Result(1, C.Total) = "計"
    BuildHeader = Result
End Function

Private Function BuildBottom(ByVal Totals As App_EnrollmentTotals, ByVal UpperColumn As Long) As Variant()
    Dim Result() As Variant
    ReDim Result(1 To 1, 1 To UpperColumn)
    Result(1, C.Grade) = "合計"
    Result(1, C.Male) = Totals.GetMainMaleGrand + Totals.GetSpecialMaleGrand
    Result(1, C.Female) = Totals.GetMainFemaleGrand + Totals.GetSpecialFemaleGrand
    Result(1, C.Total) = Totals.GetMainGradeGrand + Totals.GetSpecialGrand
    BuildBottom = Result
End Function

Private Function BuildBody(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal MainGradeCount As Long, ByVal MainClassNoCount As Long, ByVal SpecialClassNoCount As Long) As Variant()
    Dim GenderBody() As Variant
    GenderBody = BuildGenderBody(EnrollmentRMs, Totals, MainGradeCount, MainClassNoCount, SpecialClassNoCount)
    Dim ClassTotalBody() As Variant
    ClassTotalBody = BuildClassTotalBody(Totals, SpecialStreamRMs)
    BuildBody = Pre_ArrayUtility.MergeHorizontally(GenderBody, ClassTotalBody)
    
    BuildBody = Pre_ArrayUtility.MergeVertically(BuildBody, GrandTotalBody)
End Function

Private Function BuildGenderBody(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal MainGradeCount As Long, ByVal MainClassNoCount As Long, ByVal SpecialClassNoCount As Long) As Variant()
    Dim MainGenderBody() As Variant
    MainGenderBody = BuildMainGenderBody(EnrollmentRMs, Totals, MainGradeCount, MainClassNoCount)
    Dim SpecialGenderBody() As Variant
    SpecialGenderBody = BuildSpecialGenderBody(EnrollmentRMs, Totals, SpecialClassNoCount)
    BuildGenderBody = Pre_ArrayUtility.MergeVertically(MainGenderBody, SpecialGenderBody)
End Function

Private Function BuildClassTotalBody(ByVal Totals As App_EnrollmentTotals, ByVal SpecialStreamRMs As App_SpecialStreamReadModels) As Variant()
    Dim SpecialClassNoCount As Long
    SpecialClassNoCount = SpecialStreamRMs.Count
    Dim MainClassTotalBody() As Variant
    MainClassTotalBody = BuildMainClassTotalBody(Totals, SpecialClassNoCount)
    Dim SpecialClassTotalBody() As Variant
    SpecialClassTotalBody = BuildSpecialClassTotalBody(Totals, SpecialClassNoCount)
    BuildClassTotalBody = Pre_ArrayUtility.MergeVertically(MainClassTotalBody, SpecialClassTotalBody)
End Function

Private Function BuildMainGenderBody(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal GradeCount As Long, ByVal ClassNoCount As Long) As Variant()
    Dim MaleBody() As Variant
    MaleBody = BuildMainMaleBody(EnrollmentRMs, Totals, GradeCount, ClassNoCount)
    Dim FemaleBody() As Variant
    FemaleBody = BuildMainFemaleBody(EnrollmentRMs, Totals, GradeCount, ClassNoCount)
    BuildMainGenderBody = Pre_ArrayUtility.MergeHorizontally(MaleBody, FemaleBody)
End Function

Private Function BuildSpecialGenderBody(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal ClassNoCount As Long) As Variant()
    Dim MaleBody() As Variant
    MaleBody = BuildSpecialMaleBody(EnrollmentRMs, Totals, ClassNoCount)
    Dim FemaleBody() As Variant
    FemaleBody = BuildSpecialFemaleBody(EnrollmentRMs, Totals, ClassNoCount)
    BuildSpecialGenderBody = Pre_ArrayUtility.MergeHorizontally(MaleBody, FemaleBody)
End Function

Private Function BuildMainLeftItems(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal GradeCount As Long, ByVal ClassNoCount As Long) As Variant()
    Dim UpperRow As Long
    UpperRow = GradeCount * ClassNoCount + GradeCount
    Dim Result() As Variant
    ReDim Result(1 To UpperRow, 1 To 1)
    Dim i As Long
    For i = 1 To EnrollmentRMs.MainMaleCount
        Dim Model As App_EnrollmentReadModel
        Set Model = EnrollmentRMs.GetMainMale(i)
        Dim R As Long
        R = (Model.Grade - 1) * ClassNoCount + Model.ClassNo
        If 0 < (R Mod (ClassNoCount + 1)) Then
            Result(R, COLUMN_INDEX.Grade) = Model.Grade
            Result(R, COLUMN_INDEX.ClassNo) = Model.ClassNo
        Else
            Result(R, COLUMN_INDEX.Grade) = Model.Grade & "年"
            Result(R, COLUMN_INDEX.ClassNo) = "計"
        End If
    Next
    BuildMainLeftItems = Result
End Function

Private Function BuildMainMaleBody(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal GradeCount As Long, ByVal ClassNoCount As Long) As Variant()
    Dim UpperRow As Long
    UpperRow = GradeCount * ClassNoCount + GradeCount
    Dim Result() As Variant
    ReDim Result(1 To UpperRow, 1 To 1)
    Dim i As Long
    For i = 1 To EnrollmentRMs.MainMaleCount
        Dim Model As App_EnrollmentReadModel
        Set Model = EnrollmentRMs.GetMainMale(i)
        Dim R As Long
        R = (Model.Grade - 1) * ClassNoCount + Model.ClassNo
        If 0 < (R Mod (ClassNoCount + 1)) Then
            Result(R, 1) = Model.Value
        Else
            Result(R, 1) = Totals.GetMainMale(Model.Grade)
        End If
    Next
    BuildMainMaleBody = Result
End Function

Private Function BuildMainFemaleBody(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal GradeCount As Long, ByVal ClassNoCount As Long) As Variant()
    Dim UpperRow As Long
    UpperRow = GradeCount * ClassNoCount + GradeCount
    Dim Result() As Variant
    ReDim Result(1 To GradeCount, 1 To 1)
    Dim i As Long
    For i = 1 To EnrollmentRMs.MainFemaleCount
        Dim Model As App_EnrollmentReadModel
        Set Model = EnrollmentRMs.GetMainFemale(i)
        Dim R As Long
        R = (Model.Grade - 1) * ClassNoCount + Model.ClassNo
        If 0 < (R Mod (ClassNoCount + 1)) Then
            Result(R, 1) = Model.Value
        Else
            Result(R, 1) = Totals.GetMainFemale(Model.Grade)
        End If
    Next
    BuildMainFemaleBody = Result
End Function

Private Function BuildSpecialMaleBody(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal ClassNoCount As Long) As Variant()
    Dim UpperRow As Long
    UpperRow = ClassNoCount + 1
    Dim Result() As Variant
    ReDim Result(1 To UpperRow, 1 To 1)
    Dim i As Long
    For i = 1 To EnrollmentRMs.SpecialMaleCount
        Dim Model As App_EnrollmentReadModel
        Set Model = EnrollmentRMs.GetSpecialMaleGrand(i)
        Dim R As Long
        R = Model.ClassNo
        If 0 < (R Mod (ClassNoCount + 1)) Then
            Result(R, 1) = Model.Value
        Else
            Result(R, 1) = Totals.GetSpecialMaleGrand
        End If
    Next
    BuildSpecialMaleBody = Result
End Function

Private Function BuildSpecialFemaleBody(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal ClassNoCount As Long) As Variant()
    Dim UpperRow As Long
    UpperRow = ClassNoCount + 1
    Dim Result() As Variant
    ReDim Result(1 To UpperRow, 1 To 1)
    Dim i As Long
    For i = 1 To EnrollmentRMs.SpecialFemaleCount
        Dim Model As App_EnrollmentReadModel
        Set Model = EnrollmentRMs.GetSpecialFemaleGrand(i)
        Dim R As Long
        R = Model.ClassNo
        If 0 < (R Mod (ClassNoCount + 1)) Then
            Result(R, 1) = Model.Value
        Else
            Result(R, 1) = Totals.GetSpecialFemaleGrand
        End If
    Next
    BuildSpecialFemaleBody = Result
End Function

Private Function BuildMainClassTotalBody(ByVal Totals As App_EnrollmentTotals, ByVal ClassNoCount As Long) As Variant()
    Dim UpperRow As Long
    UpperRow = ClassNoCount + 1
    Dim Result() As Variant
    ReDim Result(1 To UpperRow, 1 To 1)
    Dim R As Long
    For R = 1 To UpperRow
        Dim Grade As Long
        Grade = ((R - 1) \ ClassNoCount) + 1
        Dim ClassNo As Long
        ClassNo = ((R - 1) Mod ClassNoCount) + 1
        If 0 < (R Mod (ClassNoCount + 1)) Then
            Result(R, 1) = Totals.GetMainClass(Grade, ClassNo)
        Else
            Result(R, 1) = Totals.GetMainGrade(Grade)
        End If
    Next
    BuildMainClassTotalBody = Result
End Function

Private Function BuildSpecialClassTotalBody(ByVal Totals As App_EnrollmentTotals, ByVal ClassNoCount As Long) As Variant()
    Dim UpperRow As Long
    UpperRow = ClassNoCount + 1
    Dim Result() As Variant
    ReDim Result(1 To UpperRow, 1 To 1)
    Dim R As Long
    For R = 1 To UpperRow
        Dim ClassNo As Long
        ClassNo = ((R - 1) Mod ClassNoCount) + 1
        If 0 < (R Mod (ClassNoCount + 1)) Then
            Result(R, 1) = Totals.GetSpecialClass(ClassNo)
        Else
            Result(R, 1) = Totals.GetSpecialGrand
        End If
    Next
    BuildSpecialClassTotalBody = Result
End Function

Private Function App_IDailyEnrollmentPresenter_CreateViewModel(ByVal EnrollmentRMs As App_EnrollmentReadModels, ByVal Totals As App_EnrollmentTotals, ByVal MainStreamRM As App_MainStreamReadModel, ByVal SpecialStreamRMs As App_SpecialStreamReadModels) As Pre_TableViewModel
    Dim Result As Pre_TableViewModel
    Set Result = New Pre_TableViewModel
    Result.Table = Format(EnrollmentRMs, Totals, MainStreamRM, SpecialStreamRMs)
    Set App_IDailyEnrollmentPresenter_CreateViewModel = Result
End Function
