VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregateFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal Entities As VBA.Collection) As Dom_EnrollmentAggregate
    Dim Aggregate As Dom_EnrollmentAggregate
    Set Aggregate = New Dom_EnrollmentAggregate
    Aggregate.Initialize Entities
End Function
