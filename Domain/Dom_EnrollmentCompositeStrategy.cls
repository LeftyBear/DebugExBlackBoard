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
    Dim StreamType As Long
    StreamType = ResolveStreamType(Parts(0))
    Dim Grade As Long
    Grade = ParseGrade(Parts(0))
    Dim ClassNo As Long
    ClassNo = ParseClass(Parts(1))
    Dim Gender As Long
    Gender = ParseGender(Parts(2))
    Dim GradeColumn As Dom_GradeColumn
    Set GradeColumn = New Dom_GradeColumn
    GradeColumn.Initialize Grade
    Dim ClassColumn As Dom_ClassColumn
    Set ClassColumn = New Dom_ClassColumn
    ClassColumn.Initialize ClassNo
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
    End If
End Function

Private Function ParseGrade(ByVal Part As String) As Long
    If VBA.IsNumeric(Part) Then ParseGrade = CLng(Part)
End Function

Private Function ParseClass(ByVal Part As String) As Long
    If VBA.IsNumeric(Part) Then ParseClass = CLng(Part)
End Function

Private Function ParseGender(ByVal Part As String) As Long
    Select Case Part
    Case "íj": ParseGender = Dom_Gender.Male
    Case "èó": ParseGender = Dom_Gender.Female
    End Select
End Function
