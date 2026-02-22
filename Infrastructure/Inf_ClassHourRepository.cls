VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ClassHourRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Infrastructure.Repository"
Option Explicit

Private Type Member
    Reader As Inf_IReadRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Reader As Inf_IReadRepository)
    Set This.Reader = Reader
End Sub

Public Function ReadAll(ByVal FilePath As String) As VBA.Collection
'    Dim Raw As VBA.Collection
'    Set Raw = This.Reader.ReadAll(FilePath)
'    Dim Result As VBA.Collection
'    Set Result = New VBA.Collection
'    Dim Row As Scripting.Dictionary
'    For Each Row In Raw
'        Dim Member As Dom_ClassHour
'        Set Member = MapRow(Row)
'        Result.Add Member
'    Next
'    Set ReadAll = Result
End Function

Private Function MapRow(ByVal Row As Scripting.Dictionary) As Dom_ClassHour
    Dim Member As Dom_ClassHour
    Set Member = New Dom_ClassHour
    Dim TargetDate As Date
    Dim ExeClassHourSubject() As Long
    Dim PlanClassHourSubject() As Long
    MapClassHourSubject Row, ExeClassHourSubject, PlanClassHourSubject
    Dim ExeTimeTablePeriod() As String
    Dim PlanTimeTablePeriod() As String
    MapTimeTablePeriod Row, ExeTimeTablePeriod, PlanTimeTablePeriod
'    Member.Initialize TargetDate, ExeClassHourSubject, ExeTimeTablePeriod, PlanClassHourSubject, PlanTimeTablePeriod
    Set MapRow = Member
End Function

Private Sub MapClassHourSubject(ByVal Row As Scripting.Dictionary, ByRef ExeClassHourSubject() As Long, ByRef PlanClassHourSubject() As Long)
    Dim Index As Long
    For Index = 1 To 1
        Dim Grade As Long
        For Grade = 1 To 1
            Dim Buffer As Long
            ExeClassHourSubject(Index, Grade) = Buffer
            PlanClassHourSubject(Index, Grade) = Buffer
        Next
    Next
End Sub

Private Sub MapTimeTablePeriod(ByVal Row As Scripting.Dictionary, ByRef ExeTimeTablePeriod() As String, ByRef PlanTimeTablePeriod() As String)
    Dim Index As Long
    For Index = 1 To 1
        Dim Grade As Long
        For Grade = 1 To 1
            Dim Buffer As String
            ExeTimeTablePeriod(Index, Grade) = Buffer
            PlanTimeTablePeriod(Index, Grade) = CLng(Buffer)
        Next
    Next
End Sub

