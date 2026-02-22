VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    GradeNo     As Long
    ClassNo     As Long
    GenderCode  As Long
    Count       As Long
End Type

Private This As Member

Public Sub Initialize(ByVal GradeNo As Long, ByVal ClassNo As Long, ByVal GenderCode As Long, ByVal Count As Long)
    This.GradeNo = GradeNo
    This.ClassNo = ClassNo
    This.GenderCode = GenderCode
    This.Count = Count
End Sub

Public Property Get GradeNo() As Long
    GradeNo = This.GradeNo
End Property

Public Property Get ClassNo() As Long
    ClassNo = This.ClassNo
End Property

Public Property Get GenderCode() As Long
    GenderCode = This.GenderCode
End Property

Public Property Get Count() As Long
    Count = This.Count
End Property

