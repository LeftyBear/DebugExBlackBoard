VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    Resolver    As Dom_HeaderTokenResolver
    Interpreter As Dom_ScheduleInterpreter
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_HeaderTokenResolver
    Set This.Interpreter = New Dom_ScheduleInterpreter
End Sub

Public Function Resolve(ByRef RawCSVRows() As String) As VBA.Collection
    Dim ColumnGroup As VBA.Collection
    Set ColumnGroup = New VBA.Collection
    Dim R As Long
    For R = LBound(RawCSVRows) + 1 To UBound(RawCSVRows)
        Dim Column As Dom_ScheduleColumn
        Set Column = New Dom_ScheduleColumn
        Dim ItemList As Scripting.Dictionary
        Set ItemList = New Scripting.Dictionary
        Dim C As Long
        For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
            Dim HeaderName As String
            HeaderName = RawCSVRows(1, C)
            If HeaderName = jpnHizuke Then
                Dim DateID As Dom_DateID
                DateID = This.Resolver.ResolveDate(RawCSVRows(R, C))
            Else
                Dim Schedule As Dom_Schedule
                Set Schedule = This.Interpreter.Interprete(HeaderName, RawCSVRows(R, C))
                ItemList.Add HeaderName, Schedule
            End If
        Next
        Column.Initialize DateID, ItemList
        ColumnGroup.Add Column
    Next
    Set Resolve = ColumnGroup
End Function
