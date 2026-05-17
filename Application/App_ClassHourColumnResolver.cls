VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Structure As Dom_SchoolStructure
End Type

Private This As Member

Public Sub AddOn(ByVal Structure As Dom_SchoolStructure)
    Set This.Structure = Structure
End Sub

Public Function Resolve(ByVal Headers As VBA.Collection) As Scripting.Dictionary
    Dim Map As Scripting.Dictionary
    Set Map = New Scripting.Dictionary
    Dim C As Long
    For C = 1 To Headers.Count
        Dim ColumnName As String
        ColumnName = Headers.Item(C)
        Dim Meta As App_ClassHourColumn
        Set Meta = ParseColumnMeta(ColumnName)
        Map.Add C, Meta
    Next
    Set Resolve = Map
End Function

Private Function ParseColumnMeta(ByVal ColumnName As String) As App_ClassHourColumn
    Dim Meta As App_ClassHourColumn
    Set Meta = New App_ClassHourColumn
    Dim Parts() As String
    Parts = VBA.Split(ColumnName, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnHizuke) Then
            Meta.RawDate = Part
        End If
        If 0 < VBA.InStr(1, Part, jpnKeikaku) Then
            Meta.TypeCode = Dom_ClassHourTypePolicy.Plan
        ElseIf 0 < VBA.InStr(1, Part, jpnJisshi) Then
            Meta.TypeCode = Dom_ClassHourTypePolicy.Execution
        End If
        If This.Structure.Subject.Exists(Part) Then
            Meta.Subject = Part
        ElseIf 0 < VBA.InStr(1, Part, jpnKouji) Then
            Dim PeriodString As String
            PeriodString = VBA.Replace(Part, jpnKouji, vbNullString)
            Meta.Period = Inf_NumericUtility.ExtractNumber(PeriodString)
        End If
        If 0 < VBA.InStr(1, Part, jpnNen) Then
            Dim GradeString As String
            GradeString = VBA.Replace(Part, jpnNen, vbNullString)
            Meta.Grade = Inf_NumericUtility.ExtractNumber(GradeString)
        End If
    Next
    Set ParseColumnMeta = Meta
End Function
