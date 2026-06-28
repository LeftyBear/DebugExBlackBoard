VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Formatter")
Option Explicit

Public Function Calculate(ByVal Models As App_EnrollmentReadModels, ByVal MeinKey As String, ByVal SpecialKey As String, ByVal MaleKey As String, ByVal FemaleKey As String) As App_EnrollmentTotals
    Dim Result As App_EnrollmentTotals
    Set Result = New App_EnrollmentTotals
    Dim Model As App_EnrollmentReadModel
    For Each Model In Models
        If Model.TypeKey = MeinKey Then
            If Model.Gender = MaleKey Then
                AddMainMale Model, Result
            ElseIf Model.Gender = FemaleKey Then
                AddMainFemale Model, Result
            End If
            AddMainClass Model, Result
        ElseIf Model.TypeKey = SpecialKey Then
            If Model.Gender = MaleKey Then
                AddSpecialMale Model, Result
            ElseIf Model.Gender = FemaleKey Then
                AddSpecialFemale Model, Result
            End If
            AddSpecialClass Model, Result
        End If
    Next
    Set Calculate = Result
End Function

Private Sub AddMainMale(ByVal Model As App_EnrollmentReadModel, ByVal Totals As App_EnrollmentTotals)
    Dim Total As Long
    Total = Totals.GetMainMale(Model.Grade) + Model.Value
    Totals.AddMainMale Model.Grade, Total
End Sub

Private Sub AddMainFemale(ByVal Model As App_EnrollmentReadModel, ByVal Totals As App_EnrollmentTotals)
    Dim Total As Long
    Total = Totals.GetMainFemale(Model.Grade) + Model.Value
    Totals.AddMainFemale Model.Grade, Total
End Sub

Private Sub AddMainClass(ByVal Model As App_EnrollmentReadModel, ByVal Totals As App_EnrollmentTotals)
    Dim Total As Long
    Total = Totals.GetMainClass(Model.Grade, Model.ClassNo) + Model.Value
    Totals.AddMainClass Model.Grade, Model.ClassNo, Total
End Sub

Private Sub AddSpecialMale(ByVal Model As App_EnrollmentReadModel, ByVal Totals As App_EnrollmentTotals)
    Dim Total As Long
    Total = Totals.GetSpecialMaleGrand + Model.Value
    Totals.AddSpecialMale Total
End Sub

Private Sub AddSpecialFemale(ByVal Model As App_EnrollmentReadModel, ByVal Totals As App_EnrollmentTotals)
    Dim Total As Long
    Total = Totals.GetSpecialFemaleGrand + Model.Value
    Totals.AddSpecialFemale Total
End Sub

Private Sub AddSpecialClass(ByVal Model As App_EnrollmentReadModel, ByVal Totals As App_EnrollmentTotals)
    Dim Total As Long
    Total = Totals.GetSpecialClass(Model.ClassName) + Model.Value
    Totals.AddSpecialClass Model.ClassName, Total
End Sub
