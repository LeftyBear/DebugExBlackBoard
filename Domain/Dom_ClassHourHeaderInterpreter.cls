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

Private Type Member
    ClassHourType   As String
    SubjectName     As String
    Grade           As String
End Type

Private This As Member

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Interprete(ByVal Config As Dom_SchoolConfig, ByVal ColumnDTO As App_ColumnDTO, ByRef Values() As Long)
    ParseHeader Config, ColumnDTO.HeaderName
    Dim ClassHourType As Dom_ClassHourType
    ClassHourType = ResolveClassHourType(This.ClassHourType)
    Dim Subject As Dom_Subject
    Subject = Config.ResolveSubject(This.SubjectName)
    Dim Grade As Dom_Grade
    Grade = ResolveGrade(This.Grade)
    Values(ClassHourType, Subject.Index, Grade.Value) = Dom_Function.NormalizeToLong(ColumnDTO.RawValue)
End Sub

Private Sub ParseHeader(ByVal Config As Dom_SchoolConfig, ByVal HeaderName As String)
    Dim Parts() As String
    Parts = VBA.Split(HeaderName, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        If 0 < VBA.InStr(1, Parts(i), jpnKeikaku) Or 0 < VBA.InStr(1, Parts(i), jpnJisshi) Then
            This.ClassHourType = Parts(i)
        ElseIf -1 < UBound(VBA.Filter(Config.SubjectNames, Parts(i))) Then
            This.SubjectName = Parts(i)
        ElseIf 0 < VBA.InStr(1, Parts(i), jpnNen) Then
            This.Grade = Parts(i)
        End If
    Next
End Sub

Private Function ResolveClassHourType(ByVal Part As String) As Dom_ClassHourType
    Select Case Part
    Case jpnKeikaku
        ResolveClassHourType = Dom_ClassHourType.Plan
    Case jpnJisshi
        ResolveClassHourType = Dom_ClassHourType.Execution
    End Select
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
