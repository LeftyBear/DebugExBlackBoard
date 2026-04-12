VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourRowBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    Parser      As Dom_ClassHourHeaderParser
    Resolver    As Dom_ClassHourCellResolver
End Type
Private This As Member

Private Sub Class_Initialize()
    Set This.Parser = New Dom_ClassHourHeaderParser
    Set This.Resolver = New Dom_ClassHourCellResolver
End Sub

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    This.Parser.AddOn Structure
    This.Resolver.AddOn Structure
End Sub

Public Function Build(ByRef RawCSVRows() As String, ByVal RowIndex As Long) As Dom_ClassHourAggregate
    Dim DateID As Dom_DateID
    Dim PlanValueList As New Scripting.Dictionary
    Set PlanValueList = New Scripting.Dictionary
    Dim PlanMarkList As New Scripting.Dictionary
    Set PlanMarkList = New Scripting.Dictionary
    Dim ExeValueList As New Scripting.Dictionary
    Set ExeValueList = New Scripting.Dictionary
    Dim ExeMarkList As New Scripting.Dictionary
    Set ExeMarkList = New Scripting.Dictionary
    Dim C As Long
    For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
        Dim Header As Dom_ClassHourHeader
        Set Header = This.Parser.Parse(RawCSVRows(1, C))
        This.Resolver.Resolve Header, RawCSVRows(RowIndex, C), DateID, PlanValueList, PlanMarkList, ExeValueList, ExeMarkList
    Next
    Dim Result As Dom_ClassHourAggregate
    Set Result = New Dom_ClassHourAggregate
    Result.Initialize DateID, PlanValueList, PlanMarkList, ExeValueList, ExeMarkList
    Set Build = Result
End Function
