VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ExecuteEnrollmentResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Application.DTO"
Option Explicit

Private Type Member
    Aggregate As Dom_EnrollmentAggregate
End Type

Private This As Member

Public Sub Initialize(ByVal Aggregate As Dom_EnrollmentAggregate)
    Set This.Aggregate = Aggregate
End Sub

Public Property Get Aggregate() As Dom_EnrollmentAggregate
    Set Aggregate = This.Aggregate
End Property
