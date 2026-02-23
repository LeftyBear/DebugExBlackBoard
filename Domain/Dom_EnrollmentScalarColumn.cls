VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentScalarColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit
Implements Dom_IEnrollmentColumn
Private Type Member
    HeaderName  As String
    Value       As String
End Type

Private This As Member

Public Sub Initialize(ByVal HeaderName As String, ByVal Value As String)
    If VBA.Len(HeaderName) = 0 Then Err.Raise errEmptyHeaderName, , "HeaderName is empty"
    This.HeaderName = HeaderName
    This.Value = Value
End Sub

Public Property Get HeaderName() As String
    HeaderName = This.HeaderName
End Property

Public Property Get Value() As String
    Value = This.Value
End Property

Private Property Get Dom_IEnrollmentColumn_HeaderName() As String
    Dom_IEnrollmentColumn_HeaderName = This.HeaderName
End Property

Private Property Get Dom_IEnrollmentColumn_Value() As String
    Dom_IEnrollmentColumn_Value = This.Value
End Property
