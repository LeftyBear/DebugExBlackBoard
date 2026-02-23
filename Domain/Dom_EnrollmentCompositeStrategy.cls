VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentCompositeStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Strategy")
Option Explicit

Public Function Create(ByVal ColumnDTO As App_EnrollmentColumnDTO) As Dom_IEnrollmentColumn
    Dim Parts() As String
    Parts = ParseHeader(ColumnDTO.HeaderName)
    Dim StreamType As Dom_StreamType
    StreamType = ResolveStreamType(Parts(0))
    Dim GradeColumn As Dom_GradeColumn
    Set GradeColumn = New Dom_GradeColumn
    GradeColumn.Initialize Parts(0)
    Dim ClassColumn As Dom_ClassColumn
    Set ClassColumn = New Dom_ClassColumn
    ClassColumn.Initialize Parts(1)
    Dim GenderColumn As Dom_GenderColumn
    Set GenderColumn = New Dom_GenderColumn
    GenderColumn.Initialize ParseGender(Parts(2))
    Dim CompositeColumn As Dom_EnrollmentCompositeColumn
    Set CompositeColumn = New Dom_EnrollmentCompositeColumn
    CompositeColumn.Initialize StreamType, GradeColumn, ClassColumn, GenderColumn, ColumnDTO.RawValue
    Set Create = CompositeColumn
End Function

Private Function ParseHeader(ByVal HeaderName As String) As String()
    Dim Parts() As String
    Parts = VBA.Split(HeaderName, charUnderScore)
    If UBound(Parts) <> 2 Then Err.Raise errInvalidHeaderFormat, , "Invalid composite header format"
    ParseHeader = Parts
End Function

Private Function ResolveStreamType(ByVal Part As String) As Dom_StreamType
    If Part Like "ì¡éx" Then
        ResolveStreamType = Dom_StreamType.Special
    ElseIf IsNumeric(Part) Then
        ResolveStreamType = Dom_StreamType.Main
    Else
        Err.Raise errInvalidGradeValue, , "Invalid grade value"
    End If
End Function

Private Function ParseGender(ByVal Part As String) As Dom_Gender
    Select Case Part
    Case "íj": ParseGender = Dom_Gender.Male
    Case "èó": ParseGender = Dom_Gender.Female
    End Select
End Function
