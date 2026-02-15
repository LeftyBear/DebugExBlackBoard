VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Entities As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal Entities As VBA.Collection)
    If Entities Is Nothing Then Err.Raise Dom_ErrNum.NothingEntities, , "Entities is nothing"
    Set This.Entities = Entities
    Dim Copy As VBA.Collection
    Set Copy = New VBA.Collection
    Dim Item As Variant
    For Each Item In Entities
        If Item Is Nothing Then Err.Raise Dom_ErrNum.NothingItem, , "Item is nothing"
        If Not TypeOf Item Is Dom_Enrollment Then Err.Raise Dom_ErrNum.NotEntityType, , "Item is not entity"
        Copy.Add Item
    Next
    Set This.Entities = Copy
End Sub

Public Property Get Count() As Long
    Count = This.Entities.Count
End Property

Public Function ExtractTransfer() As Dom_EnrollmentList
    Dim Filtered As VBA.Collection
    Set Filtered = New VBA.Collection
    Dim Enrollment As Dom_Enrollment
    For Each Enrollment In This.Entities
        If Enrollment.Transfer Then Filtered.Add Enrollment
    Next
    Dim EnrollmentList As Dom_EnrollmentList
    Set EnrollmentList = New Dom_EnrollmentList
    EnrollmentList.Initialize Filtered
    Set ExtractTransfer = EnrollmentList
End Function

Public Function ToCollection() As VBA.Collection
    Dim Copy As VBA.Collection
    Set Copy = New VBA.Collection
    Dim Entity As Variant
    For Each Entity In This.Entities
        Copy.Add Entity
    Next
    Set ToCollection = Copy
End Function
