VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourCellResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    ClassHourInterpreter    As Dom_ClassHourInterpreter
    TimeTableInterpreter    As Dom_TimeTableInterpreter
End Type
Private This As Member

Private Sub Class_Initialize()
    Set This.ClassHourInterpreter = New Dom_ClassHourInterpreter
    Set This.TimeTableInterpreter = New Dom_TimeTableInterpreter
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    This.ClassHourInterpreter.AddOn Structure
End Sub

'@Ignore ParameterCanBeByVal
'@Ignore ParameterCanBeByVal
'@Ignore ParameterCanBeByVal
'@Ignore ParameterCanBeByVal
Public Sub Resolve(ByVal Header As Dom_ClassHourHeader, ByVal RawText As String, ByRef DateID As Dom_DateID, ByRef PlanValueList As Scripting.Dictionary, ByRef PlanMarkList As Scripting.Dictionary, ByRef ExeValueList As Scripting.Dictionary, ByRef ExeMarkList As Scripting.Dictionary)
    If Header.RawDate <> vbNullString Then
        Set DateID = Dom_ValueObjectFactory.CreateDateID(RawText)
    ElseIf Header.ClassHourType = jpnKeikaku Then
        Dim ClassHour As Dom_ClassHour
        Set ClassHour = This.ClassHourInterpreter.Interprete(Header, RawText)
        If Not ClassHour Is Nothing Then
            If Header.SubjectName <> vbNullString Then
                PlanValueList.Add ClassHour.GetID, ClassHour
            Else
                PlanMarkList.Add ClassHour.GetID, ClassHour
            End If
        End If
    ElseIf Header.ClassHourType = jpnJisshi Then
        Dim TimeTable As Dom_TimeTable
        Set TimeTable = This.TimeTableInterpreter.Interprete(Header, RawText)
        If Not TimeTable Is Nothing Then
            If Header.SubjectName <> vbNullString Then
                ExeValueList.Add TimeTable.GetID, TimeTable
            Else
                ExeMarkList.Add TimeTable.GetID, TimeTable
            End If
        End If
    End If
End Sub
