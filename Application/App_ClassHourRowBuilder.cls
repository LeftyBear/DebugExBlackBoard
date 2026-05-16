VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourRowBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function Build(ByVal Records As VBA.Collection, ByVal Structure As Dom_SchoolStructure) As VBA.Collection
    Dim Headers As VBA.Collection
    Set Headers = Records.Item(1)
    Dim Column As App_ClassHourColumnResolver
    Set Column = New App_ClassHourColumnResolver
    Column.AddOn Structure
    Dim MetaMap As Scripting.Dictionary
    Set MetaMap = Column.Resolve(Headers)
    Dim Row As App_ClassHourRowResolver
    Set Row = New App_ClassHourRowResolver
    Set Build = Row.Resolve(Records, MetaMap)
End Function

