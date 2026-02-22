VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Model")
Option Explicit

Private Type Member
    Columns()   As Dom_EnrollmentColumn
    Count       As Long
End Type

Private This As Member

Public Sub Add(ByVal Column As Dom_EnrollmentColumn)
    This.Count = This.Count + 1
    ReDim Preserve This.Columns(1 To This.Count)
    Set This.Columns(This.Count) = Column
End Sub

Public Property Get Columns(ByRef Result() As Dom_EnrollmentColumn)
    Result = This.Columns
End Property
