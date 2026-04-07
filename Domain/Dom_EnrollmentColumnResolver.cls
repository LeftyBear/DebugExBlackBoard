VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    Parser      As Dom_EnrollmentHeaderParser
    Interpreter As Dom_EnrollmentInterpreter
    Resolver    As Dom_HeaderTokenResolver
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Parser = New Dom_EnrollmentHeaderParser
    Set This.Interpreter = New Dom_EnrollmentInterpreter
    Set This.Resolver = New Dom_HeaderTokenResolver
End Sub

Public Function Resolve(ByRef RawCSVRows() As String) As VBA.Collection
    Dim ColumnGroup As VBA.Collection
    Set ColumnGroup = New VBA.Collection
    Dim R As Long
    For R = LBound(RawCSVRows) + 1 To UBound(RawCSVRows)
        Dim ValueList As Scripting.Dictionary
        Set ValueList = New Scripting.Dictionary
        Dim MainStream As Scripting.Dictionary
        Set MainStream = New Scripting.Dictionary
        Dim SpecialStream As Scripting.Dictionary
        Set SpecialStream = New Scripting.Dictionary
        Dim C As Long
        For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
            Dim Header As Dom_EnrollmentHeader
            Set Header = This.Parser.Parse(RawCSVRows(1, C))
            If Header.RawDate <> vbNullString Then
                Dim DateID As Dom_DateID
                DateID = This.Resolver.ResolveDate(RawCSVRows(R, C))
            ElseIf Header.Transfer <> vbNullString Then
                Dim Transfer As Dom_Transfer
                Transfer = This.Resolver.ResolveTransfer(RawCSVRows(R, C))
            ElseIf Header.Remarks <> vbNullString Then
                Dim Remarks As Dom_Remarks
                Remarks = This.Resolver.ResolveRemarks(RawCSVRows(R, C))
            Else
                Dim Value As Dom_Enrollment
                Set Value = This.Interpreter.Interprete(Header, RawCSVRows(R, C))
                If Header.Grade <> jpnTokushi Then
                    MainStream.Add RawCSVRows(1, C), Value
                Else
                    SpecialStream.Add RawCSVRows(1, C), Value
                End If
            End If
        Next
        Dim Column As Dom_EnrollmentColumn
        Set Column = New Dom_EnrollmentColumn
        Column.Initialize DateID, Transfer, Remarks, MainStream, SpecialStream
        ColumnGroup.Add Column
    Next
    Set Resolve = ColumnGroup
End Function
