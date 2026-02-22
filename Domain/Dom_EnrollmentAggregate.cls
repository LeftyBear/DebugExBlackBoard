VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type Member
    Records As VBA.Collection
End Type

Private This As Member

Public Sub Initialize()
    Set This.Records = New VBA.Collection
End Sub

Public Sub Add(ByVal Record As Dom_EnrollmentRecord)
    This.Records.Add Record
End Sub

Public Property Get Records() As VBA.Collection
    Set Records = This.Records
End Property

Public Property Get Count() As Long
    Count = This.Records.Count
End Property
