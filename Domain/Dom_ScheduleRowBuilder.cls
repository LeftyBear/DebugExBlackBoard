VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleRowBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    Resolver    As Dom_ScheduleCellResolver
    Parser      As Dom_ScheduleHeaderParser
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_ScheduleCellResolver
    Set This.Parser = New Dom_ScheduleHeaderParser
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    This.Parser.AddOn Structure
End Sub

Public Function Build(ByRef RawCSVRows() As String, ByVal RowIndex As Long) As Dom_ScheduleAggregate
    Dim DateID As Dom_DateID
    Dim ItemList As Scripting.Dictionary
    Set ItemList = New Scripting.Dictionary
    Dim C As Long
    For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
        Dim Header As Dom_ScheduleHeader
        Set Header = This.Parser.Parse(RawCSVRows(1, C))
        This.Resolver.Resolve Header, RawCSVRows(RowIndex, C), DateID, ItemList
    Next
    Dim Result As Dom_ScheduleAggregate
    Set Result = New Dom_ScheduleAggregate
    Result.Initialize DateID, ItemList
    Set Build = Result
End Function
