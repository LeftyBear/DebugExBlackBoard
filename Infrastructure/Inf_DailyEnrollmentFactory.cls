VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyEnrollmentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.EntityFactory")
Option Explicit
Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Factory = New Dom_ValueObjectFactory
End Sub

Public Function Create(ByVal TargetDate As Date, ByVal Row As Inf_EnrollmentRow) As Dom_Enrollment
    Dim DateID As Dom_DateID
    Set DateID = This.Factory.CreateDateID(TargetDate)
    Dim ID As Dom_EnrollmentID
    Set ID = CreateID(Row.RawID)
    Dim Value As Dom_EnrollmentValue
    Set Value = This.Factory.CreateEnrollmentValue(Row.Value)
    If Row.Transfer <> vbNullString Then
        Dim RawTransfer As Boolean
        RawTransfer = GetTransfer(Row.Transfer)
        Dim Transfer As Dom_Transfer
        Set Transfer = This.Factory.CreateTransfer(RawTransfer)
    End If
    If Row.Remarks <> vbNullString Then
        Dim Remarks As Dom_Remarks
        Set Remarks = This.Factory.CreateRemarks(Row.Remarks)
    End If
    If ID Is Nothing And Transfer Is Nothing Then Exit Function
    If ID Is Nothing And Remarks Is Nothing Then Exit Function
    Dim Entity As Dom_Enrollment
    Set Entity = New Dom_Enrollment
    Entity.Initialize DateID, ID, Value, Transfer, Remarks
    Set Create = Entity
End Function

Private Function CreateID(ByVal RawID As String) As Dom_EnrollmentID
    If RawID = vbNullString Then Exit Function
    Dim RawTypeCode As Long
    RawTypeCode = GetTypeCode(RawID)
    Dim RawGrade As Long
    RawGrade = GetGrade(RawID)
    Dim RawClassNo As Long
    RawClassNo = GetClassNo(RawID)
    Dim RawGender As Long
    RawGender = GetGender(RawID)
    Dim StreamType As Dom_StreamType
    Set StreamType = This.Factory.CreateStreamType(RawTypeCode)
    Dim Grade As Dom_Grade
    Set Grade = This.Factory.CreateGrade(RawGrade)
    Dim ClassNo As Dom_ClassNo
    Set ClassNo = This.Factory.CreateClassNo(RawClassNo)
    Dim Gender As Dom_Gender
    Set Gender = This.Factory.CreateGender(RawGender)
    Set CreateID = This.Factory.CreateEnrollmentID(StreamType, Grade, ClassNo, Gender)
End Function

Private Function GetTypeCode(ByVal RawID As String) As Long
    Dim Parts() As String
    Parts = VBA.Split(RawID, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnNen) Then
            GetTypeCode = Dom_StreamTypePolicy.Main: Exit Function
        ElseIf 0 < VBA.InStr(1, Part, jpnTokushi) Then
            GetTypeCode = Dom_StreamTypePolicy.Special: Exit Function
        End If
    Next
End Function

Private Function GetGrade(ByVal RawID As String) As Long
    Dim Parts() As String
    Parts = VBA.Split(RawID, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnNen) Then
            Dim Buffer As String
            Buffer = VBA.Replace(Part, jpnNen, vbNullString)
            GetGrade = Inf_NumericUtility.ExtractNumber(Buffer): Exit Function
        ElseIf 0 < VBA.InStr(1, Part, jpnTokushi) Then
            GetGrade = 1: Exit Function
        End If
    Next
End Function

Private Function GetClassNo(ByVal RawID As String) As Long
    Dim Parts() As String
    Parts = VBA.Split(RawID, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnKumi) Then
            Dim Buffer As String
            Buffer = VBA.Replace(Part, jpnKumi, vbNullString)
            GetClassNo = Inf_NumericUtility.ExtractNumber(Buffer): Exit Function
        End If
    Next
End Function

Private Function GetGender(ByVal RawID As String) As Long
    Dim Parts() As String
    Parts = VBA.Split(RawID, charUnderScore)
    Dim i As Long
    For i = 0 To UBound(Parts)
        Dim Part As String
        Part = Parts(i)
        If 0 < VBA.InStr(1, Part, jpnOtoko) Then
            GetGender = Dom_GenderTypePolicy.Male: Exit Function
        ElseIf 0 < VBA.InStr(1, Part, jpnOnna) Then
            GetGender = Dom_GenderTypePolicy.Female: Exit Function
        End If
    Next
End Function

Private Function GetTransfer(ByVal RawTransfer As String) As Boolean
    GetTransfer = App_BooleanUtility.NormalizeToBoolean(RawTransfer)
End Function
