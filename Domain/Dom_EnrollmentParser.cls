VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Parser")
Option Explicit

Public Sub Parse(ByVal ColumnDTO As App_ColumnDTO, ByRef Values() As Long)
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
    Values(StreamType, Grade, ClassNo, Gender) = CLng(ColumnDTO.RawValue)
End Sub

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
