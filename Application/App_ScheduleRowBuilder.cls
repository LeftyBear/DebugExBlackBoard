VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ScheduleRowBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function Build(ByVal Records As VBA.Collection, ByVal Structure As Dom_SchoolStructure) As VBA.Collection
    Dim HeaderRow() As String
    HeaderRow = Inf_StringUtility.GetHeaderRow(Records.Item(1))
    Dim Column As App_ScheduleColumnResolver
    Set Column = New App_ScheduleColumnResolver
    Column.AddOn Structure
    Dim MetaMap As Scripting.Dictionary
    Set MetaMap = Column.Resolve(HeaderRow)
    Dim Row As App_ScheduleRowResolver
    Set Row = New App_ScheduleRowResolver
    Set Build = Row.Resolve(Records, MetaMap)
End Function
