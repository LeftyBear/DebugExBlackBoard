VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    Parser                  As Dom_ClassHourHeaderParser
    Resolver                As Dom_HeaderTokenResolver
    ClassHourInterpreter    As Dom_ClassHourInterpreter
    TimeTableInterpreter    As Dom_TimeTableInterpreter
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Parser = New Dom_ClassHourHeaderParser
    Set This.Resolver = New Dom_HeaderTokenResolver
    Set This.ClassHourInterpreter = New Dom_ClassHourInterpreter
    Set This.TimeTableInterpreter = New Dom_TimeTableInterpreter
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    This.Parser.AddOn Structure
    This.ClassHourInterpreter.AddOn Structure
End Sub

Public Function Resolve(ByRef RawCSVRows() As String) As VBA.Collection
    Dim ColumnGroup As VBA.Collection
    Set ColumnGroup = New VBA.Collection
    Dim R As Long
    For R = LBound(RawCSVRows) + 1 To UBound(RawCSVRows)
        Dim Column As Dom_ClassHourColumn
        Set Column = New Dom_ClassHourColumn
        Dim PlanValueList As Scripting.Dictionary
        Set PlanValueList = New Scripting.Dictionary
        Dim PlanMarkList As Scripting.Dictionary
        Set PlanMarkList = New Scripting.Dictionary
        Dim ExeValueList As Scripting.Dictionary
        Set ExeValueList = New Scripting.Dictionary
        Dim ExeMarkList As Scripting.Dictionary
        Set ExeMarkList = New Scripting.Dictionary
        Dim C As Long
        For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
            Dim Header As Dom_ClassHourHeader
            Set Header = This.Parser.Parse(RawCSVRows(1, C))
            If Header.RawDate <> vbNullString Then
                Dim DateID As Dom_DateID
                Set DateID = This.Resolver.ResolveDate(RawCSVRows(R, C))
            Else
                If Header.ClassHourType = jpnKeikaku Then
                    Dim ClassHour As Dom_ClassHour
                    Set ClassHour = This.ClassHourInterpreter.Interprete(Header, RawCSVRows(R, C))
                    If Not ClassHour Is Nothing Then
                        If Header.SubjectName <> vbNullString Then
                            PlanValueList.Add ClassHour.GetID, ClassHour
                        Else
                            PlanMarkList.Add ClassHour.GetID, ClassHour
                        End If
                    End If
                ElseIf Header.ClassHourType = jpnJisshi Then
                    Dim TimeTable As Dom_TimeTable
                    Set TimeTable = This.TimeTableInterpreter.Interprete(Header, RawCSVRows(R, C))
                    If Not TimeTable Is Nothing Then
                        If Header.SubjectName <> vbNullString Then
                            ExeValueList.Add TimeTable.GetID, TimeTable
                        Else
                            ExeMarkList.Add TimeTable.GetID, TimeTable
                        End If
                    End If
                End If
            End If
        Next
        Column.Initialize DateID, PlanValueList, PlanMarkList, ExeValueList, ExeMarkList
        ColumnGroup.Add Column
    Next
    Set Resolve = ColumnGroup
End Function
