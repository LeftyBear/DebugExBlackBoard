VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    FixDate     As Date
    Transfer    As Boolean
    Remarks     As String
    Values()    As Long
End Type

Private This As Member

Public Sub Initialize(ByVal FixDate As Date, ByVal Transfer As Boolean, ByVal Remarks As String, ByRef Values() As Long)
    This.FixDate = FixDate
    This.Transfer = Transfer
    This.Remarks = Remarks
    This.Values = Values
End Sub

Public Property Get FixDate() As Date
    FixDate = This.FixDate
End Property

Public Property Get Transfer() As Boolean
    Transfer = This.Transfer
End Property

Public Property Get Remarks() As String
    Remarks = This.Remarks
End Property

Public Property Get Value(ByVal StreamType As Long, ByVal Grade As Long, ByVal ClassNo As Long, ByVal Gender As Long) As Long
    Value = This.Values(StreamType, Grade, ClassNo, Gender)
End Property
