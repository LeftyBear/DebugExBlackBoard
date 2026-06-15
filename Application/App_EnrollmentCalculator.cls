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

Public Sub CalculateGenderTotal(ByVal ReadModels As App_EnrollmentReadModels, ByVal StreamTypeKeyCode As App_StreamTypeKeyCode, ByVal GenderTypeCode As Dom_GenderTypeCode, ByVal Enrollments As App_EnrollmentTotals)
    If ReadModels Is Nothing Then Exit Sub
    Dim Result As App_EnrollmentTotals
    Set Result = New App_EnrollmentTotals
    Dim StreamTypeResolver As App_StreamTypeResolver
    Set StreamTypeResolver = New App_StreamTypeResolver
    Dim StreamType As String
    StreamType = StreamTypeResolver.ParseKey(StreamTypeKeyCode)
    Dim GenderResolver As App_GenderResolver
    Set GenderResolver = New App_GenderResolver
    Dim Gender As String
    Gender = GenderResolver.ParseKey(GenderTypeCode)
    Dim FilterdModels As VBA.Collection
    Set FilterdModels = FilterByTypeAndGender(ReadModels, StreamType, Gender)
    Dim Model As App_EnrollmentReadModel
    For Each Model In FilterdModels
        Result.AddGenderValue StreamType, Gender, Model.Grade, Model.Value
    Next
End Sub

Public Sub CalculateClassTotal(ByVal ReadModels As App_EnrollmentReadModels, ByVal StreamTypeKeyCode As App_StreamTypeKeyCode, ByVal Enrollments As App_EnrollmentTotals)
    If ReadModels Is Nothing Then Exit Sub
    Dim Result As App_EnrollmentTotals
    Set Result = New App_EnrollmentTotals
    Dim StreamTypeResolver As App_StreamTypeResolver
    Set StreamTypeResolver = New App_StreamTypeResolver
    Dim StreamType As String
    StreamType = StreamTypeResolver.ParseKey(StreamTypeKeyCode)
    Dim FilterdModels As VBA.Collection
    Set FilterdModels = FilterByType(ReadModels, StreamType)
    Dim Model As App_EnrollmentReadModel
    For Each Model In FilterdModels
        Result.AddClassValue StreamType, Model.Grade, Model.ClassNo, Model.Value
    Next
End Sub

Private Function FilterByTypeAndGender(ByVal ReadModels As App_EnrollmentReadModels, ByVal StreamType As String, ByVal Gender As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Model As App_EnrollmentReadModel
    For Each Model In ReadModels
        If Model.TypeKey = StreamType Then
            If Model.Gender = Gender Then
                Result.Add Model
            End If
        End If
    Next
    Set FilterByTypeAndGender = Result
End Function

Private Function FilterByType(ByVal ReadModels As App_EnrollmentReadModels, ByVal StreamType As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Model As App_EnrollmentReadModel
    For Each Model In ReadModels
        If Model.TypeKey = StreamType Then
            Result.Add Model
        End If
    Next
    Set FilterByType = Result
End Function
