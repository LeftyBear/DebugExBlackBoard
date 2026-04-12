VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    DateID              As Dom_DateID
    Transfer            As Dom_Transfer
    Remarks             As Dom_Remarks
    MainStreamList      As Scripting.Dictionary
    SpecialStreamList   As Scripting.Dictionary
End Type

Private This As Member

Public Sub Initialize(ByVal DateID As Dom_DateID, ByVal Transfer As Dom_Transfer, ByVal Remarks As Dom_Remarks, ByVal MainStreamList As Scripting.Dictionary, ByVal SpecialStreamList As Scripting.Dictionary)
    Set This.DateID = DateID
    Set This.Transfer = Transfer
    Set This.Remarks = Remarks
    Set This.MainStreamList = MainStreamList
    Set This.SpecialStreamList = SpecialStreamList
End Sub

Public Sub UpdateTransfer(ByVal NewValue As Dom_Transfer)
    With New Dom_Transfer
        Set This.Transfer = .Create(NewValue.Value)
    End With
End Sub

Public Sub UpdateRemarks(ByVal NewValue As Dom_Remarks)
    With New Dom_Remarks
        Set This.Remarks = .Create(NewValue.Value)
    End With
End Sub

Public Sub UpdateMainStream(ByVal ID As Dom_EnrollmentID, ByVal NewValue As Dom_EnrollmentValue)
    Dim Target As Dom_Enrollment
    Set Target = FindMainStream(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_EnrollmentAggregate", "Enrollmentオブジェクトが必要です。"
    Dim Value As Dom_EnrollmentValue
    With New Dom_EnrollmentValue
        Set Value = .Create(NewValue.Value)
    End With
    Target.ChangeValueInternal Value
End Sub

Public Sub UpdateSpecialStream(ByVal ID As Dom_EnrollmentID, ByVal NewValue As Dom_EnrollmentValue)
    Dim Target As Dom_Enrollment
    Set Target = FindSpecialStream(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_EnrollmentAggregate", "Enrollmentオブジェクトが必要です。"
    Dim Value As Dom_EnrollmentValue
    With New Dom_EnrollmentValue
        Set Value = .Create(NewValue.Value)
    End With
    Target.ChangeValueInternal Value
End Sub

Public Function GetDateKey() As String
    GetDateKey = This.DateID.ToKey
End Function

Public Function GetMainStreamKeys() As Variant()
    GetMainStreamKeys = This.MainStreamList.Keys
End Function

Public Function GetSpecialStreamKeys() As Variant()
    GetSpecialStreamKeys = This.SpecialStreamList.Keys
End Function

Public Function GetMainStreams() As Variant()
    GetMainStreams = This.MainStreamList.Items
End Function

Public Function GetSpecialStreams() As Variant()
    GetSpecialStreams = This.SpecialStreamList.Items
End Function

Private Function FindMainStream(ByVal ID As Dom_EnrollmentID) As Dom_Enrollment
    Dim Key As String
    Key = ID.ToKey
    If This.MainStreamList.Exists(Key) Then
        Set FindMainStream = This.MainStreamList.Item(Key)
    End If
End Function

Private Function FindSpecialStream(ByVal ID As Dom_EnrollmentID) As Dom_Enrollment
    Dim Key As String
    Key = ID.ToKey
    If This.SpecialStreamList.Exists(Key) Then
        Set FindSpecialStream = This.SpecialStreamList.Item(Key)
    End If
End Function
