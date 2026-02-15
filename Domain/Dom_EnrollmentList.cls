VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.EntityList"
Option Explicit

Private Type Member
    Items As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal Items As VBA.Collection)
    Set This.Items = Items
End Sub

Public Property Get Items() As VBA.Collection
    Set Items = This.Items
End Property

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Public Function FilterTransfer() As Dom_EnrollmentList
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Item As Dom_Enrollment
    For Each Item In This.Items
        If Item.Transfer Then Result.Add Item
    Next
    Dim List As Dom_EnrollmentList
    Set List = New Dom_EnrollmentList
    List.Initialize Result
    Set FilterTransfer = List
End Function

Public Function SearchRemarksByKeywords(ByVal Keywords As VBA.Collection) As Dom_EnrollmentList
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Item As Dom_Enrollment
    For Each Item In This.Items
        If RemarksMatchAll(Item.Remarks, Keywords) Then Result.Add Item
    Next
    Dim List As Dom_EnrollmentList
    Set List = New Dom_EnrollmentList
    List.Initialize Result
    Set SearchRemarksByKeywords = List
End Function

Private Function RemarksMatchAll(ByVal Remarks As String, ByVal Keywords As VBA.Collection) As Boolean
    Dim Keyword As Variant
    For Each Keyword In Keywords
        Dim Word As String
        Word = VBA.Trim$(VBA.CStr(Keyword))
        If 0 < VBA.Len(Word) Then
            If VBA.InStr(1, Remarks, Word, VBA.vbTextCompare) = 0 Then
                RemarksMatchAll = False
                Exit Function
            End If
        End If
    Next
    RemarksMatchAll = True
End Function

