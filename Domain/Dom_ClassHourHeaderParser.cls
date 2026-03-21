VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourHeaderParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Parser")
Option Explicit

Public Function Parse(ByVal Config As Dom_SchoolStructure, ByVal HeaderName As String) As Dom_ClassHourHeader
    Dim Parts() As String
    Parts = VBA.Split(HeaderName, charUnderScore)
    If UBound(Parts) = 2 Then
        Dim i As Long
        For i = 0 To UBound(Parts)
            If 0 < VBA.InStr(1, Parts(i), jpnKeikaku) Or 0 < VBA.InStr(1, Parts(i), jpnJisshi) Then
                Dim ClassHourType As String
                ClassHourType = Parts(i)
            ElseIf Config.IsSubject(Parts(i)) Then
                Dim SubjectName As String
                SubjectName = Parts(i)
            ElseIf 0 < VBA.InStr(1, Parts(i), jpnKouji) Then
                Dim Period As String
                Period = Parts(i)
            ElseIf 0 < VBA.InStr(1, Parts(i), jpnNen) Then
                Dim Grade As String
                Grade = Parts(i)
            End If
        Next
    Else
        Dim TargetDate As String
        TargetDate = Parts(0)
    End If
    Dim Result As Dom_ClassHourHeader
    Set Result = New Dom_ClassHourHeader
    Result.Initialize TargetDate, ClassHourType, SubjectName, Period, Grade
    Set Parse = Result
End Function
