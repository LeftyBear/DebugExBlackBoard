VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.DTOCollection")
Option Explicit
Private Type Member
    GenderTotals    As VBA.Collection
    ClassTotals     As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.GenderTotals = New VBA.Collection
    Set This.ClassTotals = New VBA.Collection
End Sub

Public Sub AddGenderValue(ByVal TypeKey As String, ByVal Gender As String, ByVal Grade As Long, ByVal Value As Long)
    Dim Total As Long
    Total = Me.GenderTotalItem(TypeKey, Gender, Grade) + Value
    Set This.GenderTotals.Item(TypeKey & Gender & CStr(Grade)) = Total
End Sub

Public Sub AddClassValue(ByVal TypeKey As String, ByVal Grade As Long, ByVal ClassNo As Long, ByVal Value As Long)
    Dim Total As Long
    Total = Me.ClassTotalItem(TypeKey, Grade, ClassNo) + Value
    Set This.ClassTotals.Item(TypeKey & CStr(Grade) & CStr(ClassNo)) = Total
End Sub

Public Function GenderTotalCount() As Long
    GenderTotalCount = This.GenderTotals.Count
End Function

Public Function ClassTotalCount() As Long
    ClassTotalCount = This.ClassTotals.Count
End Function

Public Function GenderTotalItem(ByVal TypeKey As String, ByVal Gender As String, ByVal Grade As Long) As Long
    If This.GenderTotals.Item(TypeKey & Gender & CStr(Grade)) Is Nothing Then Exit Function
    GenderTotalItem = This.GenderTotals.Item(TypeKey & Gender & CStr(Grade))
End Function

Public Function ClassTotalItem(ByVal TypeKey As String, ByVal Grade As Long, ByVal ClassNo As Long) As Long
    If This.ClassTotals.Item(TypeKey & CStr(Grade) & CStr(ClassNo)) Is Nothing Then Exit Function
    ClassTotalItem = This.ClassTotals.Item(TypeKey & CStr(Grade) & CStr(ClassNo))
End Function
