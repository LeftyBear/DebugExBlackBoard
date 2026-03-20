VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    TargetDate  As String
'    StreamType  As String
    Grade       As String
    ClassNo     As String
    Gender      As String
    Transfer    As String
    Remarks     As String
End Type

Private This As Member

Public Sub Initialize(ByVal TargetDate As String, ByVal Grade As String, ByVal ClassNo As String, ByVal Gender As String, ByVal Transfer As String, ByVal Remarks As String)
    This.TargetDate = TargetDate
'    This.StreamType = StreamType
    This.Grade = Grade
    This.ClassNo = ClassNo
    This.Gender = Gender
    This.Transfer = Transfer
    This.Remarks = Remarks
End Sub

Public Property Get TargetDate() As String
    TargetDate = This.TargetDate
End Property
'
'Public Property Get StreamType() As String
'    StreamType = This.StreamType
'End Property

Public Property Get Grade() As String
    Grade = This.Grade
End Property

Public Property Get ClassNo() As String
    ClassNo = This.ClassNo
End Property

Public Property Get Gender() As String
    Gender = This.Gender
End Property

Public Property Get Transfer() As String
    Transfer = This.Transfer
End Property

Public Property Get Remarks() As String
    Remarks = This.Remarks
End Property
