VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentRowBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    Resolver    As Dom_EnrollmentCellResolver
    Parser      As Dom_EnrollmentHeaderParser
End Type
Private This As Member

Private Sub Class_Initialize()
    Set This.Resolver = New Dom_EnrollmentCellResolver
    Set This.Parser = New Dom_EnrollmentHeaderParser
End Sub

Public Function Build(ByRef RawCSVRows() As String, ByVal RowIndex As Long) As Dom_EnrollmentAggregate
    Dim DateID As Dom_DateID
    Dim Transfer As Dom_Transfer
    Dim Remarks As Dom_Remarks
    Dim MainStreamList As New Scripting.Dictionary
    Set MainStreamList = New Scripting.Dictionary
    Dim SpecialStreamList As New Scripting.Dictionary
    Set SpecialStreamList = New Scripting.Dictionary
    Dim C As Long
    For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
        Dim Header As Dom_EnrollmentHeader
        Set Header = This.Parser.Parse(RawCSVRows(1, C))
        This.Resolver.Resolve Header, RawCSVRows(RowIndex, C), DateID, Transfer, Remarks, MainStreamList, SpecialStreamList
    Next
    Dim Result As Dom_EnrollmentAggregate
    Set Result = New Dom_EnrollmentAggregate
    Result.Initialize DateID, Transfer, Remarks, MainStreamList, SpecialStreamList
    Set Build = Result
End Function
