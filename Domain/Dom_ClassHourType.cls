VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Value As Dom_ClassHourTypeCode
End Type

Private This As Member

Public Sub Initialize(ByVal Value As Dom_ClassHourTypeCode)
    This.Value = Value
End Sub

Public Property Get Value() As Dom_ClassHourTypeCode
    Value = This.Value
End Property

Public Function IsPlan() As Boolean
    IsPlan = (This.Value = PlanCode)
End Function

Public Function IsExecution() As Boolean
    IsExecution = (This.Value = ExeCode)
End Function

Public Function Equals(ByVal Other As Dom_ClassHourType) As Boolean
    Equals = (This.Value = Other.Value)
End Function
