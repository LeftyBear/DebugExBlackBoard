VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentHeaderParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Parser")
Option Explicit

Public Function Parse(ByVal HeaderName As String) As Dom_EnrollmentHeader
    Dim Parts() As String
    Parts = VBA.Split(HeaderName, charUnderScore)
    If UBound(Parts) = 2 Then
        Dim i As Long
        For i = 0 To UBound(Parts)
            If 0 < VBA.InStr(1, Parts(i), jpnTokushi) Or 0 < VBA.InStr(1, Parts(i), jpnNen) Then
                Dim Grade As String
                Grade = Parts(i)
            ElseIf 0 < VBA.InStr(1, Parts(i), jpnKumi) Then
                Dim ClassNo As String
                ClassNo = Parts(i)
            ElseIf 0 < VBA.InStr(1, Parts(i), jpnOtoko) Or 0 < VBA.InStr(1, Parts(i), jpnOnna) Then
                Dim Gender As String
                Gender = Parts(i)
            End If
        Next
    ElseIf 0 < VBA.InStr(1, Parts(i), jpnHizuke) Then
        Dim TargetDate As String
        TargetDate = Parts(0)
    ElseIf 0 < VBA.InStr(1, Parts(i), jpnIdou) Then
        Dim Transfer As String
        Transfer = Parts(0)
    ElseIf 0 < VBA.InStr(1, Parts(i), jpnBikou) Then
        Dim Remarks As String
        Remarks = Parts(0)
    End If
    Dim Result As Dom_EnrollmentHeader
    Set Result = New Dom_EnrollmentHeader
    Result.Initialize TargetDate, Grade, ClassNo, Gender, Transfer, Remarks
    Set Parse = Result
End Function
