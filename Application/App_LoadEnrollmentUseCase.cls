VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_LoadEnrollmentUseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Repository  As Dom_IEnrollmentRepository
End Type

Private This As Member

Public Sub Initialize(ByVal Repository As Dom_IEnrollmentRepository)
    Set This.Repository = Repository
End Sub

Public Function Execute(ByVal SchoolYear As Long) As Dom_IEnrollmentRepository
'    Dim Records As VBA.Collection
'    Set Records = This.Repository.ReadAll(SchoolYear)
'    If Records.Count = 0 Then Exit Function
'    Dim Rows As VBA.Collection
'    Set Rows = This.Builder.Build(Records)
'    Dim Yearly As VBA.Collection
'    Set Yearly = This.Factory.Create(Rows)
'    Dim i As Long
'    For i = 1 To Yearly.Count
'        Dim Daily As Dom_DailyEnrollment
'        Set Daily = Yearly.Item(i)
'        This.Repository.Add Daily
'    Next
'    Set Execute = This.Repository
End Function
