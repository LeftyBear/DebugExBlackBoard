VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Calculator")
Option Explicit

Public Function Calculate(ByRef Items() As Variant, ByVal Filter As Dom_EnrollmentFilter) As Long
    Dim i As Long
    For i = LBound(Items) To UBound(Items)
        Dim Item As Dom_Enrollment
        Set Item = Items(i)
        If Filter.IsMatch(Item) Then
            Dim Total As Long
            Total = Total + Item.GetValue.Value
        End If
    Next
    Calculate = Total
End Function
