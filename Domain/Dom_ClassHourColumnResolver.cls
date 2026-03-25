VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Private Type Member
    Parser      As Dom_ClassHourHeaderParser
    Interpreter As Dom_ClassHourInterpreter
End Type

Private This As Member

Public Sub Initialize(ByVal Structure As Dom_SchoolStructure)
    Set This.Parser = New Dom_ClassHourHeaderParser
    This.Parser.Initialize Structure
    Set This.Interpreter = New Dom_ClassHourInterpreter
    This.Interpreter.Initialize Structure
End Sub

Public Function Resolve(ByRef RawCSVRows() As String) As VBA.Collection
    Dim EntityGroups As VBA.Collection
    Set EntityGroups = New VBA.Collection
    Dim R As Long
    For R = LBound(RawCSVRows) + 1 To UBound(RawCSVRows)
        Dim EntityGroup As VBA.Collection
        Set EntityGroup = New VBA.Collection
        Dim C As Long
        For C = LBound(RawCSVRows, 2) To UBound(RawCSVRows, 2)
            Dim Header As Dom_ClassHourHeader
            Set Header = This.Parser.Parse(RawCSVRows(1, C))
            Dim ClassHour As Dom_ClassHour
            Set ClassHour = This.Interpreter.Interprete(Header, RawCSVRows(R, C))
            EntityGroup.Add ClassHour
        Next
        EntityGroups.Add EntityGroup
    Next
    Set Resolve = EntityGroups
End Function
