VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Execution As Dom_ExecutionHourSummary
    Plan As Dom_PlanHourSummary
    IsInitialized As Boolean
End Type

Private This As Member

Friend Sub Initialize(ByVal Execution As Dom_ExecutionHourSummary, ByVal Plan As Dom_PlanHourSummary)
    If This.IsInitialized Then Err.Raise Dom_ErrNum.InitializedClass, , "Class is Initialized"
    Set This.Execution = Execution
    Set This.Plan = Plan
    This.IsInitialized = True
End Sub

Public Property Get Execution() As Dom_ExecutionHourSummary
    Set Execution = This.Execution
End Property

Public Property Get Plan() As Dom_PlanHourSummary
    Set Plan = This.Plan
End Property

