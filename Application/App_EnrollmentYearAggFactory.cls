VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentYearAggFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    EntityFactory As Dom_EnrollmentFactory
End Type

Private This As Member

Public Sub Initialize(ByVal EntityFactory As Dom_EnrollmentFactory)
    Set This.EntityFactory = EntityFactory
End Sub

Public Function Create(ByVal Rows As VBA.Collection) As Dom_EnrollmentYearAggregate
    Dim Yearly As Dom_EnrollmentYearAggregate
    Set Yearly = New Dom_EnrollmentYearAggregate
    Dim Grouped As Scripting.Dictionary
    Set Grouped = GroupByDate(Rows)
    Dim Key As Variant
    For Each Key In Grouped.Keys
        Dim Daily As Dom_EnrollmentAggregate
        Set Daily = CreateDaily(Grouped.Item(Key))
        Yearly.Add Daily
    Next
    Set Create = Yearly
End Function

Private Function GroupByDate(ByVal Rows As VBA.Collection) As Scripting.Dictionary
    Dim Groups As Scripting.Dictionary
    Set Groups = New Scripting.Dictionary
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As App_EnrollmentRow
        Set Row = Rows.Item(i)
        Dim Key As String
        Key = CStr(Row.NormDate)
        If Not Groups.Exists(Key) Then
            Groups.Add Key, New VBA.Collection
        End If
        Dim Items As VBA.Collection
        Set Items = Groups.Item(Key)
        Items.Add Row
    Next
    Set GroupByDate = Groups
End Function

Private Function CreateDaily(ByVal Rows As VBA.Collection) As Dom_EnrollmentAggregate
    Dim Aggregate As Dom_EnrollmentAggregate
    Set Aggregate = New Dom_EnrollmentAggregate
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As App_EnrollmentRow
        Set Row = Rows.Item(i)
        Dim Entity As Dom_Enrollment
        Set Entity = This.EntityFactory.Create(Row)
        If Not Entity Is Nothing Then Aggregate.Add Entity
    Next
    Set CreateDaily = Aggregate
End Function
