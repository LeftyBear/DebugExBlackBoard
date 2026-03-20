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
    If Token = jpnTokushi Then
        ResolveStreamType = Dom_StreamType.Special
    ElseIf 0 < VBA.InStr(1, Token, jpnNen) Then
        ResolveStreamType = Dom_StreamType.Main
    Else
        ResolveStreamType = 0
    End If
End Function

Public Function ResolveClassHourType(ByVal Token As String) As Dom_ClassHourType
    If Token = jpnKeikaku Then
        ResolveClassHourType = Dom_ClassHourType.Plan
    ElseIf Token = jpnJisshi Then
        ResolveClassHourType = Dom_ClassHourType.Execution
    Else
        ResolveClassHourType = 0
    End If
End Function

Public Function ResolveGender(ByVal Token As String) As Dom_Gender
    If Token = jpnOtoko Then
        ResolveGender = Dom_Gender.Male
    ElseIf Token = jpnOnna Then
        ResolveGender = Dom_Gender.Female
    Else
        ResolveGender = 0
    End If
End Function

Public Function ResolveGrade(ByVal Token As String) As Dom_Grade
    Dim TemporaryValue As String
    TemporaryValue = VBA.Replace(Token, jpnNen, vbNullString)
    If VBA.IsNumeric(TemporaryValue) Then
        Dim Value As Long
        Value = Dom_Util_Long.NormalizeToLong(TemporaryValue)
    Else
        Value = 0
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
    Else
        Value = 0
    End If
    With New Dom_ClassNo
        Set ResolveClassNo = .Create(Value)
    End With
End Function

Public Function ResolvePeriod(ByVal Part As String) As Dom_Period
    Dim TemporaryValue As String
    TemporaryValue = VBA.Replace(Part, jpnKouji, vbNullString)
    If VBA.IsNumeric(TemporaryValue) Then
        Dim Value As Long
        Value = Dom_Util_Long.NormalizeToLong(TemporaryValue)
    Else
        Value = 0
    End If
    With New Dom_Period
        Set ResolvePeriod = .Create(Value)
    End With
End Function
