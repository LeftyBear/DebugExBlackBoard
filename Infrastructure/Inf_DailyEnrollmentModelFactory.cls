VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyEnrollmentModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal DailyRows As Inf_EnrollmentRows) As App_EnrollmentReadModels
    Dim Result As App_EnrollmentReadModels
    Set Result = New App_EnrollmentReadModels
    Dim TypeMapper As Inf_StreamTypeMapper
    Set TypeMapper = New Inf_StreamTypeMapper
    Dim GradeMapper As Inf_GradeMapper
    Set GradeMapper = New Inf_GradeMapper
    Dim ClassNoMapper As Inf_ClassNoMapper
    Set ClassNoMapper = New Inf_ClassNoMapper
    Dim GenderMapper As Inf_GenderMapper
    Set GenderMapper = New Inf_GenderMapper
    Dim i As Long
    For i = 1 To DailyRows.Count
        Dim Row As Inf_EnrollmentRow
        Set Row = DailyRows.Item(i)
        Dim Model As App_EnrollmentReadModel
        Set Model = New App_EnrollmentReadModel
        If Row.RawID <> vbNullString Then
            Model.TypeCode = TypeMapper.ToDomain(Row.RawID)
            Model.Grade = ParseGrade(Row.RawID)
            Model.ClassNo = ParseClassNo(Row.RawID)
            Model.Gender = ParseGender(Row.RawID)
            Model.Value = Row.Value
        ElseIf Row.RawTransfer <> vbNullString Then
            Model.Transfer = App_BooleanUtility.NormalizeToBoolean(Row.RawTransfer)
        ElseIf Row.Remarks <> vbNullString Then
            Model.Remarks = Row.Remarks
        End If
        Result.Add Model
    Next
    Set Create = Result
End Function

Private Function ParseTypeCode(ByVal RawID As String) As Long
    Dim Parts() As String
    Parts = VBA.Split(RawID, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnNen) Then
            ParseTypeCode = Dom_StreamTypePolicy.Main: Exit Function
        ElseIf 0 < VBA.InStr(1, Part, jpnTokushi) Then
            ParseTypeCode = Dom_StreamTypePolicy.Special: Exit Function
        End If
    Next
End Function

Private Function ParseGrade(ByVal RawID As String) As Long
    Dim Parts() As String
    Parts = VBA.Split(RawID, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnNen) Then
            Dim Buffer As String
            Buffer = VBA.Replace(Part, jpnNen, vbNullString)
            ParseGrade = Inf_NumericUtility.ExtractNumber(Buffer): Exit Function
        ElseIf 0 < VBA.InStr(1, Part, jpnTokushi) Then
            ParseGrade = 1: Exit Function
        End If
    Next
End Function

Private Function ParseClassNo(ByVal RawID As String) As Long
    Dim Parts() As String
    Parts = VBA.Split(RawID, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnKumi) Then
            Dim Buffer As String
            Buffer = VBA.Replace(Part, jpnKumi, vbNullString)
            ParseClassNo = Inf_NumericUtility.ExtractNumber(Buffer): Exit Function
        End If
    Next
End Function

Private Function ParseGender(ByVal RawID As String) As Long
    Dim Parts() As String
    Parts = VBA.Split(RawID, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnOtoko) Then
            ParseGender = Dom_GenderTypePolicy.Male: Exit Function
        ElseIf 0 < VBA.InStr(1, Part, jpnOnna) Then
            ParseGender = Dom_GenderTypePolicy.Female: Exit Function
        End If
    Next
End Function

Private Function ParseTransfer(ByVal RawTransfer As String) As Boolean
    ParseTransfer = App_BooleanUtility.NormalizeToBoolean(RawTransfer)
End Function
