VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_CSVDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.DTO")
Option Explicit

Private Type Member
    Source() As String
End Type

Private This As Member

Public Sub Initialize(ByRef Values() As String)
    This.Source = Values
End Sub

Public Function RowCount() As Long
    If Not Not This.Source Then
        RowCount = UBound(This.Source) - LBound(This.Source) + 1
    End If
End Function

Public Function ColumnCount() As Long
    If Not Not This.Source Then
        ColumnCount = UBound(This.Source, 2) - LBound(This.Source, 2) + 1
    End If
End Function

Public Function HeaderName(ByVal Index As Long) As String
    If Not Not This.Source Then
        HeaderName = This.Source(LBound(This.Source), Index)
    End If
End Function

Public Function Value(ByVal RowIndex As Long, ByVal ColumnIndex As Long) As String
    If Not Not This.Source Then
        Value = This.Source(RowIndex, ColumnIndex)
    End If
End Function
