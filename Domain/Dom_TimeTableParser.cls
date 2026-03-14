VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Parser")
Option Explicit

Private Enum Index
    ClassHourType
    Period
    Grade
End Enum

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Parse(ByVal ColumnDTO As App_ColumnDTO, ByRef TimeTables() As String)
    Dim Parts() As String
    Parts = ParseHeader(ColumnDTO.HeaderName)
    Dim ClassHourType As Long
    ClassHourType = ResolveClassHourType(Parts(Index.ClassHourType))
    Dim Period As Long
    Period = ParsePeriod(Parts(Index.Period))
    Dim Grade As Long
    Grade = ParseGrade(Parts(Index.Grade))
    TimeTables(ClassHourType, Period, Grade) = ColumnDTO.RawValue
End Sub

Private Function ParseHeader(ByVal HeaderName As String) As String()
    Dim Buffer As String
    Buffer = VBA.Replace(VBA.Replace(HeaderName, jpnNen, vbNullString), jpnKouji, vbNullString)
    Dim Parts() As String
    Parts = VBA.Split(Buffer, charUnderScore)
    ParseHeader = Parts
End Function

Private Function ResolveClassHourType(ByVal Part As String) As Long
    Select Case Part
    Case jpnKeikaku
        ResolveClassHourType = Dom_ClassHourType.Plan
    Case jpnJisshi
        ResolveClassHourType = Dom_ClassHourType.Execution
    End Select
End Function

Private Function ParsePeriod(ByVal Part As String) As Long
    If VBA.IsNumeric(Part) Then ParsePeriod = CLng(Part)
End Function

Private Function ParseGrade(ByVal Part As String) As Long
    If VBA.IsNumeric(Part) Then ParseGrade = CLng(Part)
End Function
