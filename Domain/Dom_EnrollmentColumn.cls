VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    NormDate        As Dom_DateID
    Transfer        As Dom_Transfer
    Remarks         As Dom_Remarks
    MainStream      As Scripting.Dictionary
    SpecialStream   As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal NormDate As Dom_DateID, ByVal Transfer As Dom_Transfer, ByVal Remarks As Dom_Remarks, ByVal MainStream As Scripting.Dictionary, ByVal SpecialStream As Scripting.Dictionary)
    Set This.NormDate = NormDate
    Set This.Transfer = Transfer
    Set This.Remarks = Remarks
    Set This.MainStream = MainStream
    Set This.SpecialStream = SpecialStream
 End Sub

Friend Sub UpdateTransfer(ByVal NewTransfer As Boolean)
    With New Dom_Transfer
        Dim Value As Dom_Transfer
        Set Value = .Create(NewTransfer)
    End With
    Set This.Transfer = Value
End Sub

Friend Sub UpdateRemarks(ByVal NewRemarks As String)
    With New Dom_Remarks
        Dim Value As Dom_Remarks
        Set Value = .Create(NewRemarks)
    End With
    Set This.Remarks = Value
End Sub

Public Sub UpdateMainStream(ByVal ID As Dom_EnrollmentID, ByVal NewValue As Long)
    Dim Enrollment As Dom_Enrollment
    Set Enrollment = GetMainStreamByID(ID)
    If Enrollment Is Nothing Then Err.Raise DomErrNothingObject, "Dom_EnrollmentColumn", "IDをキーに持つオブジェクトが必要です。"
    With New Dom_EnrollmentValue
        Dim Value As Dom_EnrollmentValue
        Set Value = .Create(NewValue)
    End With
    Enrollment.ChangeValueInternal Value
End Sub

Public Sub UpdateSpecialStream(ByVal ID As Dom_EnrollmentID, ByVal NewValue As Long)
    Dim Enrollment As Dom_Enrollment
    Set Enrollment = GetSpecialStreamByID(ID)
    If Enrollment Is Nothing Then Err.Raise DomErrNothingObject, "Dom_EnrollmentColumn", "IDをキーに持つオブジェクトが必要です。"
    With New Dom_EnrollmentValue
        Dim Value As Dom_EnrollmentValue
        Set Value = .Create(NewValue)
    End With
    Enrollment.ChangeValueInternal Value
End Sub

Public Function GetDateKey() As String
    GetDateKey = This.NormDate.ToKey
End Function

Public Function GetTransfer() As Boolean
    GetTransfer = This.Transfer.Value
End Function

Public Function GetRemarks() As String
    GetRemarks = This.Remarks.Value
End Function

Public Function GetMainStreamValue(ByVal ID As Dom_EnrollmentID) As Long
    Dim Enrollment As Dom_Enrollment
    Set Enrollment = GetMainStreamByID(ID)
    If Enrollment Is Nothing Then Err.Raise DomErrNothingObject, "Dom_EnrollmentColumn", "IDをキーに持つオブジェクトが必要です。"
    GetMainStreamValue = Enrollment.GetValue.Value
End Function

Public Function GetSpecialStreamValue(ByVal ID As Dom_EnrollmentID) As Long
    Dim Enrollment As Dom_Enrollment
    Set Enrollment = GetSpecialStreamByID(ID)
    If Enrollment Is Nothing Then Err.Raise DomErrNothingObject, "Dom_EnrollmentColumn", "IDをキーに持つオブジェクトが必要です。"
    GetSpecialStreamValue = Enrollment.GetValue.Value
End Function

Public Function GetTotal(ByVal Filter As Dom_EnrollmentFilter) As Long
    '@Ignore SelfAssignedDeclaration
    Dim Calculator As New Dom_TimeTableMarkCalculator
    GetTotal = Calculator.Calculate(EnrollmentItems(Filter), Filter)
End Function

Private Function GetMainStreamByID(ByVal ID As Dom_EnrollmentID) As Dom_Enrollment
    Dim Key As String
    Key = ID.ToKey
    If This.MainStream.Exists(Key) Then
        Set GetMainStreamByID = This.MainStream.Item(Key)
    End If
End Function

Private Function GetSpecialStreamByID(ByVal ID As Dom_EnrollmentID) As Dom_Enrollment
    Dim Key As String
    Key = ID.ToKey
    If This.SpecialStream.Exists(Key) Then
        Set GetSpecialStreamByID = This.SpecialStream.Item(Key)
    End If
End Function

Private Function EnrollmentItems(ByVal Filter As Dom_EnrollmentFilter) As Variant()
    If Not Filter.HasType And Not Filter.HasClassNo And Not Filter.HasGender Then
        EnrollmentItems = This.MainStream.Items
    ElseIf Not Filter.HasGrade And Not Filter.HasClassNo And Not Filter.HasGender Then
        EnrollmentItems = This.SpecialStream.Items
    End If
End Function
