VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_TimeTableMarkCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.UseCase")
Option Explicit

Public Function Calculate(ByRef Items() As Variant, ByVal Filter As Dom_TimeTableFilter) As Long
    Dim i As Long
    For i = LBound(Items) To UBound(Items)
        Dim Item As Dom_TimeTable
        Set Item = Items(i)
        If Filter.IsMatch(Item) Then
            If Item.GetMark <> vbNullString Then
                Dim Count As Long
                Count = Count + 1
            End If
        End If
    Next
    Calculate = Count
End Function
