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

Public Sub CalculateGenderTotal(ByVal ReadModels As App_EnrollmentReadModels, ByVal StreamTypeKey As String, ByVal GenderKey As String, ByVal Totals As App_EnrollmentTotals)
    If ReadModels Is Nothing Then Exit Sub
    Dim GenderResolver As App_GenderResolver
    Set GenderResolver = New App_GenderResolver
    Dim FilterdModels As VBA.Collection
    Set FilterdModels = FilterByTypeAndGender(ReadModels, StreamTypeKey, GenderKey)
    Dim Model As App_EnrollmentReadModel
    For Each Model In FilterdModels
        Totals.AddGenderValue StreamTypeKey, GenderKey, Model.Grade, Model.Value
    Next
End Sub

Public Sub CalculateClassTotal(ByVal ReadModels As App_EnrollmentReadModels, ByVal StreamTypeKey As String, ByVal Totals As App_EnrollmentTotals)
    If ReadModels Is Nothing Then Exit Sub
    Dim FilterdModels As VBA.Collection
    Set FilterdModels = FilterByType(ReadModels, StreamTypeKey)
    Dim Model As App_EnrollmentReadModel
    For Each Model In FilterdModels
        Totals.AddClassValue StreamTypeKey, Model.Grade, Model.ClassNo, Model.Value
    Next
End Sub

Private Function FilterByTypeAndGender(ByVal ReadModels As App_EnrollmentReadModels, ByVal StreamTypeKey As String, ByVal GenderKey As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Model As App_EnrollmentReadModel
    For Each Model In ReadModels
        If Model.TypeKey = StreamTypeKey Then
            If Model.Gender = GenderKey Then
                Result.Add Model
            End If
        End If
    Next
    Set FilterByTypeAndGender = Result
End Function

Private Function FilterByType(ByVal ReadModels As App_EnrollmentReadModels, ByVal StreamTypeKey As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Model As App_EnrollmentReadModel
    For Each Model In ReadModels
        If Model.TypeKey = StreamTypeKey Then
            Result.Add Model
        End If
    Next
    Set FilterByType = Result
End Function
