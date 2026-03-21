VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_HeaderTokenResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Resolver")
Option Explicit

Public Function ResolveStreamType(ByVal Token As String) As Dom_StreamType
    Dim Value As Long
    If 0 < VBA.InStr(1, Token, jpnNen) Then
        Value = Dom_StreamTypeEnum.Main
    ElseIf Token = jpnTokushi Then
        Value = Dom_StreamTypeEnum.Special
    End If
    With New Dom_StreamType
        Set ResolveStreamType = .Create(Value)
    End With
End Function

Public Function ResolveClassHourType(ByVal Token As String) As Dom_ClassHourType
    Dim Value As Long
    Select Case Token
    Case jpnKeikaku
        Value = Dom_ClassHourTypeEnum.Plan
    Case jpnJisshi
        Value = Dom_ClassHourTypeEnum.Execution
    End Select
    With New Dom_ClassHourType
        Set ResolveClassHourType = .Create(Value)
    End With
End Function

Public Function ResolveGrade(ByVal Token As String) As Dom_Grade
    Dim TemporaryValue As String
    TemporaryValue = VBA.Replace(Token, jpnNen, vbNullString)
    If VBA.IsNumeric(TemporaryValue) Then
        Dim Value As Long
        Value = Dom_Util_Long.NormalizeToLong(TemporaryValue)
    End If
    With New Dom_Grade
        Set ResolveGrade = .Create(Value)
    End With
End Function

Public Function ResolveClassNo(ByVal Token As String) As Dom_ClassNo
    Dim TemporaryValue As String
    TemporaryValue = VBA.Replace(Token, jpnKumi, vbNullString)
    If VBA.IsNumeric(TemporaryValue) Then
        Dim Value As Long
        Value = Dom_Util_Long.NormalizeToLong(TemporaryValue)
    End If
    With New Dom_ClassNo
        Set ResolveClassNo = .Create(Value)
    End With
End Function

Public Function ResolveGender(ByVal Token As String) As Dom_Gender
    Dim Value As Long
    Select Case Token
    Case jpnOtoko
        Value = Dom_GenderEnum.Male
    Case jpnOnna
        Value = Dom_GenderEnum.Female
    End Select
    With New Dom_Gender
        Set ResolveGender = .Create(Value)
    End With
End Function

Public Function ResolvePeriod(ByVal Token As String) As Dom_Period
    Dim TemporaryValue As String
    TemporaryValue = VBA.Replace(Token, jpnKouji, vbNullString)
    If VBA.IsNumeric(TemporaryValue) Then
        Dim Value As Long
        Value = Dom_Util_Long.NormalizeToLong(TemporaryValue)
    End If
    With New Dom_Period
        Set ResolvePeriod = .Create(Value)
    End With
End Function

Public Function ResolveTransfer(ByVal Token As String) As Boolean
    ResolveTransfer = Dom_Util_Boolean.NormalizeToBoolean(Token)
End Function

Public Function ResolveRemarks(ByVal Token As String) As String
    ResolveRemarks = Token
End Function
