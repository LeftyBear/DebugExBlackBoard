VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Entities        As VBA.Collection
    IsInitialized   As Boolean
End Type

Private This As Member

Public Sub Initialize(ByVal Entities As VBA.Collection)
    If This.IsInitialized Then Err.Raise Dom_ErrNum.InitializedClass, , "Class is initialized"
    If Entities Is Nothing Then Err.Raise Dom_ErrNum.NothingEntities, , "Entities is nothing"
    If Entities.Count = 0 Then Err.Raise Dom_ErrNum.NothingEntities, , "Entities is nothing"
    Dim Copy As VBA.Collection
    Set Copy = New VBA.Collection
    Dim Entity As Dom_Enrollment
    For Each Entity In This.Entities
        Copy.Add Entity
    Next
    Set This.Entities = Copy
    This.IsInitialized = True
End Sub

Public Function Count() As Long
    Count = This.Entities.Count
End Function

Public Function CreateSummary() As Dom_EnrollmentSummary
    Dim Main As Dom_MainStreamSummary
    Set Main = New Dom_MainStreamSummary
    Main.Build This.Entities
    Dim Special As Dom_SpecialSummary
    Set Special = New Dom_SpecialSummary
    Special.Build This.Entities
    Dim Summary As Dom_EnrollmentSummary
    Set Summary = New Dom_EnrollmentSummary
    Summary.Initialize Main, Special
    Set CreateSummary = Summary
End Function

Public Function ExtractTransfer() As Dom_EnrollmentAggregate
    Dim Filtered As VBA.Collection
    Set Filtered = New VBA.Collection
    Dim Enrollment As Dom_Enrollment
    For Each Enrollment In This.Entities
        If Enrollment.Transfer Then Filtered.Add Enrollment
    Next
    Dim Aggregate As Dom_EnrollmentAggregate
    Set Aggregate = New Dom_EnrollmentAggregate
    Aggregate.Initialize Filtered
    Set ExtractTransfer = Aggregate
End Function
