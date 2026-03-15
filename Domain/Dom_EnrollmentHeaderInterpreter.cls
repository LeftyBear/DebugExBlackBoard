VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentHeaderInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Parser")
Option Explicit

Private Type Member
    StreamType  As String
    ClassNo     As String
    Gender      As String
End Type

Private This As Member

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Interprete(ByVal ColumnDTO As App_ColumnDTO, ByRef Values() As Long)
    ParseHeader ColumnDTO.HeaderName
    Dim StreamType As Dom_StreamType
    StreamType = ResolveStreamType(This.StreamType)
    Dim Grade As Dom_Grade
    Grade = ResolveGrade(This.StreamType)
    Dim ClassNo As Dom_ClassNo
    ClassNo = ResolveClassNo(This.ClassNo)
    Dim Gender As Dom_Gender
    Gender = ResolveGender(This.Gender)
    Values(StreamType, Grade.Value, ClassNo.Value, Gender) = Dom_Function.NormalizeToLong(ColumnDTO.RawValue)
End Sub

Private Sub ParseHeader(ByVal HeaderName As String)
    Dim Parts() As String
    Parts = VBA.Split(HeaderName, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        If 0 < VBA.InStr(1, Parts(i), jpnNen) Or 0 < VBA.InStr(1, Parts(i), jpnTokushi) Then
            This.StreamType = Parts(i)
        ElseIf 0 < VBA.InStr(1, Parts(i), jpnKumi) Then
            This.ClassNo = Parts(i)
        ElseIf 0 < VBA.InStr(1, Parts(i), jpnOtoko) Or 0 < VBA.InStr(1, Parts(i), jpnOnna) Then
            This.Gender = Parts(i)
        End If
    Next
End Sub

Private Function ResolveStreamType(ByVal Part As String) As Dom_StreamType
    If Part Like jpnTokushi Then
        ResolveStreamType = Dom_StreamType.Special
    ElseIf VBA.IsNumeric(Part) Then
        ResolveStreamType = Dom_StreamType.Main
    End If
End Function

Private Function ResolveGrade(ByVal Part As String) As Dom_Grade
    Dim Result As Dom_Grade
    Set Result = New Dom_Grade
    Dim Buffer As String
    Buffer = VBA.Replace(Part, jpnNen, vbNullString)
    If VBA.IsNumeric(Buffer) Then
        Result.Initialize CLng(Buffer)
    Else
        Result.Initialize 0
    End If
    Set ResolveGrade = Result
End Function

Private Function ResolveClassNo(ByVal Part As String) As Dom_ClassNo
    Dim Result As Dom_ClassNo
    Set Result = New Dom_ClassNo
    Dim Buffer As String
    Buffer = VBA.Replace(Part, jpnKumi, vbNullString)
    If VBA.IsNumeric(Buffer) Then
        Result.Initialize CLng(Buffer)
    Else
        Result.Initialize 0
    End If
    Set ResolveClassNo = Result
End Function

Private Function ResolveGender(ByVal Part As String) As Dom_Gender
    Select Case Part
    Case jpnOtoko: ResolveGender = Dom_Gender.Male
    Case jpnOnna: ResolveGender = Dom_Gender.Female
    End Select
End Function
