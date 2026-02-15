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
    Config As App_ClassHourConfig
End Type
Private This As Member

Public Sub Initialize(ByVal Reader As Inf_IReadRepository, ByVal Config As App_ClassHourConfig)
    Set This.Reader = Reader
    Set This.Config = Config
End Sub

Public Function ReadAll(ByVal FilePath As String) As VBA.Collection
    Dim Raw As VBA.Collection
    Set Raw = This.Reader.ReadAll(FilePath)
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Row As Scripting.Dictionary
    For Each Row In Raw
        Dim Member As Dom_ClassHour
        Set Member = MapRow(Row)
        Result.Add Member
    Next
    Set ReadAll = Result
End Function

Private Function MapRow(ByVal Row As Scripting.Dictionary) As Dom_ClassHour
    Dim Member As Dom_ClassHour
    Set Member = New Dom_ClassHour
    Dim TargetDate As Date
    TargetDate = Util_Function.NormalizeToDate(Row.Item(This.Config.DateColumnName))
    Dim ExeClassHourSubject() As Long
    Dim PlanClassHourSubject() As Long
    MapClassHourSubject Row, ExeClassHourSubject, PlanClassHourSubject
    Dim ExeTimeTablePeriod() As String
    Dim PlanTimeTablePeriod() As String
    MapTimeTablePeriod Row, ExeTimeTablePeriod, PlanTimeTablePeriod
    Member.Inject TargetDate, ExeClassHourSubject, ExeTimeTablePeriod, PlanClassHourSubject, PlanTimeTablePeriod
    Set MapRow = Member
End Function

Private Sub MapClassHourSubject(ByVal Row As Scripting.Dictionary, ByRef ExeClassHourSubject() As Long, ByRef PlanClassHourSubject() As Long)
    ReDim ExeClassHourSubject(1 To This.Config.MaxSubject, 1 To This.Config.MaxGrade)
    ReDim PlanClassHourSubject(1 To This.Config.MaxSubject, 1 To This.Config.MaxGrade)
    Dim Index As Long
    For Index = 1 To This.Config.MaxSubject
        Dim Grade As Long
        For Grade = 1 To This.Config.MaxGrade
            Dim Buffer As Long
            Buffer = Util_Function.NormalizeToLong(Row.Item(This.Config.ExeClassHourColumnName(Grade, "Subject", Index)))
            ExeClassHourSubject(Index, Grade) = Buffer
            Buffer = Util_Function.NormalizeToLong(Row.Item(This.Config.PlanClassHourColumnName(Grade, "Subject", Index)))
            PlanClassHourSubject(Index, Grade) = Buffer
        Next
    Next
End Sub

Private Sub MapTimeTablePeriod(ByVal Row As Scripting.Dictionary, ByRef ExeTimeTablePeriod() As String, ByRef PlanTimeTablePeriod() As String)
    ReDim ExeTimeTablePeriod(1 To This.Config.MaxPeriod, 1 To This.Config.MaxGrade)
    ReDim PlanTimeTablePeriod(1 To This.Config.MaxPeriod, 1 To This.Config.MaxGrade)
    Dim Index As Long
    For Index = 1 To This.Config.MaxPeriod
        Dim Grade As Long
        For Grade = 1 To This.Config.MaxGrade
            Dim Buffer As String
            Buffer = Row.Item(This.Config.ExeTimeTableColumnName(Grade, "Period", Index))
            ExeTimeTablePeriod(Index, Grade) = Buffer
            Buffer = Row.Item(This.Config.PlanTimeTableColumnName(Grade, "Period", Index))
            PlanTimeTablePeriod(Index, Grade) = CLng(Buffer)
        Next
    Next
End Sub
