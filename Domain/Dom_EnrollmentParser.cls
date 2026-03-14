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

Private Enum Index
    StreamType
    ClassNo
    Gender
End Enum

Public Sub Parse(ByVal ColumnDTO As App_ColumnDTO, ByRef Values() As Long)
    Dim Parts() As String
    Parts = ParseHeader(ColumnDTO.HeaderName)
    Dim StreamType As Long
    StreamType = ResolveStreamType(Parts(Index.StreamType))
    Dim Grade As Long
    Grade = ParseGrade(Parts(Index.StreamType))
    Dim ClassNo As Long
    ClassNo = ParseClass(Parts(Index.ClassNo))
    Dim Gender As Long
    Gender = ParseGender(Parts(Index.Gender))
    Values(StreamType, Grade, ClassNo, Gender) = Dom_Function.NormalizeToLong(ColumnDTO.RawValue)
End Sub

Private Function ParseHeader(ByVal HeaderName As String) As String()
    Dim Buffer As String
    Buffer = VBA.Replace(VBA.Replace(HeaderName, jpnNen, vbNullString), jpnKumi, vbNullString)
    Dim Parts() As String
    Parts = VBA.Split(Buffer, charUnderScore)
    ParseHeader = Parts
End Function

Private Function ResolveStreamType(ByVal Part As String) As Long
    If Part Like jpnTokushi Then
        ResolveStreamType = Dom_StreamType.Special
    ElseIf VBA.IsNumeric(Part) Then
        ResolveStreamType = Dom_StreamType.Main
    End If
End Function

Private Function ParseGrade(ByVal Part As String) As Long
    If VBA.IsNumeric(Part) Then
        ParseGrade = CLng(Part)
    ElseIf Part Like jpnTokushi Then
        ParseGrade = 1
    End If
End Function

Private Function ParseClass(ByVal Part As String) As Long
    If VBA.IsNumeric(Part) Then ParseClass = CLng(Part)
End Function

Private Function ParseGender(ByVal Part As String) As Long
    Select Case Part
    Case jpnOtoko: ParseGender = Dom_Gender.Male
    Case jpnOnna: ParseGender = Dom_Gender.Female
    End Select
End Function
