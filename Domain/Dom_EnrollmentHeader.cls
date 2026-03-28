VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.DTO")
Option Explicit

Private Type Member
    RawDate     As String
    Grade       As String
    ClassNo     As String
    Gender      As String
    Transfer    As String
    Remarks     As String
End Type

Private This As Member

Public Sub Initialize(ByVal RawDate As String, ByVal Grade As String, ByVal ClassNo As String, ByVal Gender As String, ByVal Transfer As String, ByVal Remarks As String)
    This.RawDate = RawDate
    This.Grade = Grade
    This.ClassNo = ClassNo
    This.Gender = Gender
    This.Transfer = Transfer
    This.Remarks = Remarks
End Sub

Public Property Get RawDate() As String
    RawDate = This.RawDate
End Property

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
