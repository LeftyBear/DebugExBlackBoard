VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentSummaryFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Formatter")
Option Explicit

Public Sub FormatMainStream(ByVal Summary As Dom_EnrollmentSummary, ByVal LimitValue As Dom_LimitValue, ByRef OutputArray() As Long)
    ReDim OutputArray(0 To Dom_StreamType.Special, 1 To LimitValue.MainStreamGradeCount, 1 To LimitValue.MainStreamClassCount, 0 To 1)
    Dim Grade As Long
    For Grade = 1 To LimitValue.MainStreamGradeCount
        Dim ClassNo As Long
        For ClassNo = 1 To LimitValue.MainStreamClassCount
            Dim Gender As Long
            For Gender = Dom_Gender.Male To Dom_Gender.Female
                OutputArray(Dom_StreamType.Main, Grade, ClassNo, Gender) = Summary.GetTotal(Dom_StreamType.Main, Grade, ClassNo, Gender)
            Next
        Next
    Next
End Sub

Public Sub FormatSpecial(ByVal Summary As Dom_EnrollmentSummary, ByVal LimitValue As Dom_LimitValue, ByRef OutputArray() As Long)
    ReDim OutputArray(0 To Dom_StreamType.Special, 1 To LimitValue.MainStreamGradeCount, 1 To LimitValue.MainStreamClassCount, 0 To 1)
    Dim ClassNo As Long
    For ClassNo = 1 To LimitValue.MainStreamClassCount
        Dim Gender As Long
        For Gender = Dom_Gender.Male To Dom_Gender.Female
            OutputArray(Dom_StreamType.Special, 1, ClassNo, Gender) = Summary.GetTotal(Dom_StreamType.Special, 1, ClassNo, Gender)
        Next
    Next
End Sub
