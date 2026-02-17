VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type TMember
    MainStream As Dom_MainStreamSummary
    Special As Dom_SpecialSummary
    IsInitialized As Boolean
End Type

Private This As TMember

Friend Sub Initialize(ByVal Main As Dom_MainStreamSummary, ByVal Special As Dom_SpecialSummary)
    If This.IsInitialized Then Err.Raise Dom_ErrNum.InitializedClass, , "Class is Initialized"
    Set This.MainStream = Main
    Set This.Special = Special
    This.IsInitialized = True
End Sub

Public Property Get MainStream() As Dom_MainStreamSummary
    Set MainStream = This.MainStream
End Property

Public Property Get Special() As Dom_SpecialSummary
    Set Special = This.Special
End Property

