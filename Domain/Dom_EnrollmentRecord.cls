VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    StreamType  As Dom_StreamType
    GradeNo     As Long
    ClassNo     As Long
    GenderCode  As Dom_Gender
    Count       As Long
End Type

Private This As Member

Public Sub Initialize(ByVal StreamType As Dom_StreamType, ByVal GradeNo As Long, ByVal ClassNo As Long, ByVal GenderCode As Dom_Gender, ByVal Count As Long)
    This.StreamType = StreamType
    This.GradeNo = GradeNo
    This.ClassNo = ClassNo
    This.GenderCode = GenderCode
    This.Count = Count
End Sub

Public Property Get StreamType() As Dom_StreamType
    StreamType = This.StreamType
End Property

Public Property Get GradeNo() As Long
    GradeNo = This.GradeNo
End Property

Public Property Get ClassNo() As Long
    ClassNo = This.ClassNo
End Property

Public Property Get GenderCode() As Dom_Gender
    GenderCode = This.GenderCode
End Property

Public Property Get Count() As Long
    Count = This.Count
End Property
