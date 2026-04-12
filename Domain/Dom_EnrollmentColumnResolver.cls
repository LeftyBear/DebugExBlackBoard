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
    Builder As Dom_EnrollmentRowBuilder
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Builder = New Dom_EnrollmentRowBuilder
End Sub

Public Function Resolve(ByRef RawCSVRows() As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim R As Long
    For R = LBound(RawCSVRows) + 1 To UBound(RawCSVRows)
        Dim Column As Dom_EnrollmentAggregate
        Set Column = This.Builder.Build(RawCSVRows, R)
        Result.Add Column
    Next
    Set Resolve = Result
End Function
