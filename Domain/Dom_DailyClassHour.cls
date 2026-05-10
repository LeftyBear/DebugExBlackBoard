VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DailyClassHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    DateID      As Dom_DateID
    ClassHours  As VBA.Collection
    TimeTables  As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.ClassHours = New VBA.Collection
    Set This.TimeTables = New VBA.Collection
End Sub

Public Sub AddClassHour(ByVal Entity As Dom_ClassHour)
    If This.DateID Is Nothing Then Set This.DateID = Entity.DateID
    This.ClassHours.Add Entity
End Sub

Public Sub AddTimeTable(ByVal Entity As Dom_TimeTable)
    If This.DateID Is Nothing Then Set This.DateID = Entity.DateID
    This.TimeTables.Add Entity
End Sub

Public Sub UpdateClassHourValue(ByVal ID As Dom_ClassHourID, ByVal NewValue As Long)
    Dim Target As Dom_ClassHour
    Set Target = FindClassHour(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_DailyClassHour", "ClassHourオブジェクトが必要です。"
    Dim VO As Dom_ClassHourValue
    Set VO = New Dom_ClassHourValue
    VO.Initialize NewValue
    Target.ChangeValueInternal VO
End Sub

Public Sub UpdateTimeTableValue(ByVal ID As Dom_TimeTableID, ByVal NewValue As String)
    Dim Target As Dom_TimeTable
    Set Target = FindTimeTable(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_DailyClassHour", "ClassHourオブジェクトが必要です。"
    Dim VO As Dom_TimeTableValue
    Set VO = New Dom_TimeTableValue
    VO.Initialize NewValue
    Target.ChangeValueInternal VO
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Function ClassHoursCount() As Long
    ClassHoursCount = This.ClassHours.Count
End Function

Public Function TimeTablesCount() As Long
    TimeTablesCount = This.TimeTables.Count
End Function

Public Function ClassHoursItem(ByVal Index As Long) As Dom_ClassHour
    Set ClassHoursItem = This.ClassHours.Item(Index)
End Function

Public Function TimeTablesItem(ByVal Index As Long) As Dom_TimeTable
    Set TimeTablesItem = This.TimeTables.Item(Index)
End Function

Private Function FindClassHour(ByVal ID As Dom_ClassHourID) As Dom_ClassHour
    Dim i As Long
    For i = 1 To This.ClassHours.Count
        Dim ClassHour As Dom_ClassHour
        Set ClassHour = This.ClassHours.Item(i)
        If ClassHour.ID.Equals(ID) Then
            Set FindClassHour = ClassHour: Exit Function
        End If
    Next
End Function

Private Function FindTimeTable(ByVal ID As Dom_TimeTableID) As Dom_TimeTable
    Dim i As Long
    For i = 1 To This.TimeTables.Count
        Dim TimeTable As Dom_TimeTable
        Set TimeTable = This.TimeTables.Item(i)
        If TimeTable.ID.Equals(ID) Then
            Set FindTimeTable = TimeTable: Exit Function
        End If
    Next
End Function
