VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourHeaderInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Parser")
Option Explicit

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Interprete(ByVal Config As Dom_SchoolConfig, ByVal ColumnDTO As App_ColumnDTO, ByRef Values() As Long)
    Dim Header As Dom_ClassHourHeader
    Set Header = ParseHeader(ColumnDTO.HeaderName)
    Dim ClassHourType As Dom_ClassHourType
    ClassHourType = ResolveClassHourType(Header.ClassHourType)
    Dim Subject As Dom_Subject
    Set Subject = Config.ResolveSubject(Header.SubjectName)
    If Subject Is Nothing Then Err.Raise DomErrInvalidRange, "Dom_ClassHourHeaderInterpreter", "集計名は設定値である必要があります。"
    Dim Grade As Dom_Grade
    Set Grade = ResolveGrade(Header.Grade)
    If Config.MainStreamGradeCount < Grade.Value Then Err.Raise DomErrInvalidRange, "Dom_ClassHourHeaderInterpreter", "学年は設定範囲内である必要があります。"
    Values(ClassHourType, Subject.Index, Grade.Value) = Dom_Function.NormalizeToLong(ColumnDTO.RawValue)
End Sub

Private Function ParseHeader(ByVal HeaderName As String) As Dom_ClassHourHeader
    Dim Parts() As String
    Parts = VBA.Split(HeaderName, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        If 0 < VBA.InStr(1, Parts(i), jpnKeikaku) Or 0 < VBA.InStr(1, Parts(i), jpnJisshi) Then
            Dim ClassHourType As String
            ClassHourType = Parts(i)
        ElseIf 0 < VBA.InStr(1, Parts(i), jpnNen) Then
            Dim Grade As String
            Grade = Parts(i)
        Else
            Dim SubjectName As String
            SubjectName = Parts(i)
        End If
    Next
    Dim Result As Dom_ClassHourHeader
    Set Result = New Dom_ClassHourHeader
    Result.Initialize ClassHourType, SubjectName, Grade
    Set ParseHeader = Result
End Function

Private Function ResolveClassHourType(ByVal Part As String) As Dom_ClassHourType
    Select Case Part
    Case jpnKeikaku
        ResolveClassHourType = Dom_ClassHourType.Plan
    Case jpnJisshi
        ResolveClassHourType = Dom_ClassHourType.Execution
    End Select
End Function

Private Function ResolveGrade(ByVal Part As String) As Dom_Grade
    Dim Grade As Dom_Grade
    Set Grade = New Dom_Grade
    Dim Buffer As String
    Buffer = VBA.Replace(Part, jpnNen, vbNullString)
    If VBA.IsNumeric(Buffer) Then
        Set ResolveGrade = Grade.Create(CLng(Buffer))
    Else
        Set ResolveGrade = Grade.Create(0)
    End If
End Function
