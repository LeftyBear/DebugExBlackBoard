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

Public Function Create(ByVal ColumnDTO As App_ColumnDTO) As Dom_IEntityColumn
    Dim Parts() As String
    Parts = ParseHeader(ColumnDTO.HeaderName)
    If UBound(Parts) <> 2 Then Exit Function
'    If UBound(Parts) <> 2 Then Err.Raise DomErrInvalidNaming, , "Invalid naming of header"
    Dim StreamType As Long
    StreamType = ResolveStreamType(Parts(0))
    If StreamType = 0 Then Exit Function
    Dim GradeColumn As Dom_GradeColumn
    Set GradeColumn = New Dom_GradeColumn
    GradeColumn.Initialize Parts(0)
    If Not VBA.IsNumeric(Parts(1)) Then Exit Function
    Dim ClassColumn As Dom_ClassColumn
    Set ClassColumn = New Dom_ClassColumn
    ClassColumn.Initialize Parts(1)
    Dim Gender As Long
    Gender = ParseGender(Parts(2))
    If Gender = 0 Then Exit Function
    Dim GenderColumn As Dom_GenderColumn
    Set GenderColumn = New Dom_GenderColumn
    GenderColumn.Initialize Gender
    Dim CompositeColumn As Dom_EnrollmentCompositeColumn
    Set CompositeColumn = New Dom_EnrollmentCompositeColumn
    CompositeColumn.Initialize StreamType, GradeColumn, ClassColumn, GenderColumn, ColumnDTO.RawValue
    Set Create = CompositeColumn
End Function

Private Function ParseHeader(ByVal HeaderName As String) As String()
    Dim Buffer As String
    Buffer = VBA.Replace(VBA.Replace(HeaderName, "îN", vbNullString), "ëg", vbNullString)
    Dim Parts() As String
    Parts = VBA.Split(Buffer, charUnderScore)
    ParseHeader = Parts
End Function

Private Function ResolveStreamType(ByVal Part As String) As Long
    If Part Like "ì¡éx" Then
        ResolveStreamType = Dom_StreamType.Special
    ElseIf VBA.IsNumeric(Part) Then
        ResolveStreamType = Dom_StreamType.Main
'    Else
'        Err.Raise DomErrInvalidValue, , "Invalid grade value"
    End If
End Function

Private Function ParseGender(ByVal Part As String) As Long
    Select Case Part
    Case "íj": ParseGender = Dom_Gender.Male
    Case "èó": ParseGender = Dom_Gender.Female
    End Select
End Function
