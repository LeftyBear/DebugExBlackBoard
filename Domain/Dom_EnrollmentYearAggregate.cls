VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentYearAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Aggregate")
Option Explicit

Private Type MainStreamMember
    ClassTotal()    As Long
    MaleTotal()     As Long
    FemaleTotal()   As Long
    GradeTotal()    As Long
End Type

Private Type SpecialStreamMember
    ClassTotal()    As Long
    MaleTotal()     As Long
    FemaleTotal()   As Long
    GradeTotal()    As Long
End Type

Private Type Member
    Map             As Scripting.Dictionary
    Main            As MainStreamMember
    Special         As SpecialStreamMember
    GrandTotal()    As Long
End Type

Private This As Member

Public Sub Initialize(ByVal StreamCount As Long, ByVal GradeCount As Long, ByVal ClassNoCount As Long, ByVal GenderCount As Long)
    Set This.Map = New Scripting.Dictionary
    ReDim This.Main.ClassTotal(366, StreamCount, GradeCount, ClassNoCount)
    ReDim This.Main.MaleTotal(366, StreamCount, GradeCount)
    ReDim This.Main.FemaleTotal(366, StreamCount, GradeCount)
    ReDim This.Main.GradeTotal(366, StreamCount, GradeCount)
    ReDim This.Special.ClassTotal(366, StreamCount, GradeCount, ClassNoCount)
    ReDim This.Special.MaleTotal(366, StreamCount, GradeCount)
    ReDim This.Special.FemaleTotal(366, StreamCount, GradeCount)
    ReDim This.Special.GradeTotal(366, StreamCount, GradeCount)
    ReDim This.GrandTotal(366)
End Sub

Public Sub Add(ByVal TargetDate As Date, ByVal Aggregate As Dom_EnrollmentAggregate)
    This.Map.Add TargetDate, Aggregate
    SumValue TargetDate, Aggregate
End Sub

Public Sub UpdateHours(ByVal TargetDate As Date, ByVal StreamType As Dom_StreamType, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo, ByVal Gender As Dom_Gender, ByVal NewValue As Long)
    Dim Aggregate As Dom_EnrollmentAggregate
    Aggregate = This.Map.Item(TargetDate)
    Dim Diff As Long
    Diff = Aggregate.UpdateValue(StreamType, Grade, ClassNo, Gender, NewValue)
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    Select Case StreamType
    Case Dom_StreamType.Main
        This.Main.ClassTotal(DayCount, StreamType, Grade.Value, ClassNo.Value) = This.Main.ClassTotal(DayCount, StreamType, Grade.Value, ClassNo.Value) + Diff
        This.Main.MaleTotal(DayCount, StreamType, Grade.Value) = This.Main.MaleTotal(DayCount, StreamType, Grade.Value) + Diff
        This.Main.FemaleTotal(DayCount, StreamType, Grade.Value) = This.Main.FemaleTotal(DayCount, StreamType, Grade.Value) + Diff
        This.Main.GradeTotal(DayCount, StreamType, Grade.Value) = This.Main.GradeTotal(DayCount, StreamType, Grade.Value) + Diff
    Case Dom_StreamType.Special
        This.Special.ClassTotal(DayCount, StreamType, Grade.Value, ClassNo.Value) = This.Special.ClassTotal(DayCount, StreamType, Grade.Value, ClassNo.Value) + Diff
        This.Special.MaleTotal(DayCount, StreamType, Grade.Value) = This.Special.MaleTotal(DayCount, StreamType, Grade.Value) + Diff
        This.Special.FemaleTotal(DayCount, StreamType, Grade.Value) = This.Special.FemaleTotal(DayCount, StreamType, Grade.Value) + Diff
        This.Special.GradeTotal(DayCount, StreamType, Grade.Value) = This.Special.GradeTotal(DayCount, StreamType, Grade.Value) + Diff
    End Select
    This.GrandTotal(DayCount) = This.GrandTotal(DayCount) + Diff
End Sub

Public Function GetAggregate(ByVal TargetDate As Date) As Dom_EnrollmentAggregate
    Set GetAggregate = This.Map.Item(TargetDate)
End Function

Public Function GetMainClassTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetMainClassTotal = This.Main.ClassTotal(DayCount, Dom_StreamType.Main, Grade.Value, ClassNo.Value)
End Function

Public Function GetMainMaleTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetMainMaleTotal = This.Main.MaleTotal(DayCount, Dom_StreamType.Main, Grade.Value)
End Function

Public Function GetMainFemaleTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetMainFemaleTotal = This.Main.FemaleTotal(DayCount, Dom_StreamType.Main, Grade.Value)
End Function

Public Function GetMainGradeTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetMainGradeTotal = This.Main.GradeTotal(DayCount, Dom_StreamType.Main, Grade.Value)
End Function

Public Function GetSpecialClassTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade, ByVal ClassNo As Dom_ClassNo) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetSpecialClassTotal = This.Special.ClassTotal(DayCount, Dom_StreamType.Special, Grade.Value, ClassNo.Value)
End Function

Public Function GetSpecialMaleTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetSpecialMaleTotal = This.Special.MaleTotal(DayCount, Dom_StreamType.Special, Grade.Value)
End Function

Public Function GetSpecialFemaleTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetSpecialFemaleTotal = This.Special.FemaleTotal(DayCount, Dom_StreamType.Special, Grade.Value)
End Function

Public Function GetSpecialGradeTotal(ByVal TargetDate As Date, ByVal Grade As Dom_Grade) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetSpecialGradeTotal = This.Special.GradeTotal(DayCount, Dom_StreamType.Special, Grade.Value)
End Function

Public Function GetGrandTotal(ByVal TargetDate As Date) As Long
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    GetGrandTotal = This.GrandTotal(DayCount)
End Function

Private Sub SumValue(ByVal TargetDate As Date, ByVal Aggregate As Dom_EnrollmentAggregate)
    Dim DayCount As Long
    DayCount = Dom_Util_Date.DiffApr1st(TargetDate)
    Dim G As Long
    For G = 1 To Aggregate.GradeCount
        Dim Grade As Dom_Grade
        Set Grade = CreateGrade(G)
        Dim C As Long
        For C = 1 To Aggregate.ClassNoCount
            Dim ClassNo As Dom_ClassNo
            Set ClassNo = CreateClassNo(C)
            Dim MainMale As Long, MainFemale As Long, SpecialMale As Long, SpecialFemale As Long
            MainMale = Aggregate.GetValue(Dom_StreamType.Main, Grade, ClassNo, Dom_Gender.Male)
            MainFemale = Aggregate.GetValue(Dom_StreamType.Main, Grade, ClassNo, Dom_Gender.Female)
            SpecialMale = Aggregate.GetValue(Dom_StreamType.Special, Grade, ClassNo, Dom_Gender.Male)
            SpecialFemale = Aggregate.GetValue(Dom_StreamType.Special, Grade, ClassNo, Dom_Gender.Female)
            This.Main.ClassTotal(DayCount, Dom_StreamType.Main, G, C) = This.Main.ClassTotal(DayCount, Dom_StreamType.Main, G, C) + MainMale + MainFemale
            This.Main.MaleTotal(DayCount, Dom_StreamType.Main, G) = This.Main.MaleTotal(DayCount, Dom_StreamType.Main, G) + MainMale
            This.Main.FemaleTotal(DayCount, Dom_StreamType.Main, G) = This.Main.FemaleTotal(DayCount, Dom_StreamType.Main, G) + MainFemale
            This.Main.GradeTotal(DayCount, Dom_StreamType.Main, G) = This.Main.GradeTotal(DayCount, Dom_StreamType.Main, G) + MainMale + MainFemale
            This.Special.ClassTotal(DayCount, Dom_StreamType.Main, G, C) = This.Special.ClassTotal(DayCount, Dom_StreamType.Main, G, C) + MainMale + MainFemale
            This.Special.MaleTotal(DayCount, Dom_StreamType.Main, G) = This.Special.MaleTotal(DayCount, Dom_StreamType.Main, G) + MainMale
            This.Special.FemaleTotal(DayCount, Dom_StreamType.Main, G) = This.Special.FemaleTotal(DayCount, Dom_StreamType.Main, G) + MainFemale
            This.Special.GradeTotal(DayCount, Dom_StreamType.Main, G) = This.Special.GradeTotal(DayCount, Dom_StreamType.Main, G) + MainMale + MainFemale
            This.GrandTotal(DayCount) = This.GrandTotal(DayCount) + MainMale + MainFemale + SpecialMale + SpecialFemale
        Next
    Next
End Sub

Private Function CreateGrade(ByVal Value As Long) As Dom_Grade
    With New Dom_Grade
        Set CreateGrade = .Create(Value)
    End With
End Function

Private Function CreateClassNo(ByVal Value As Long) As Dom_ClassNo
    With New Dom_ClassNo
        Set CreateClassNo = .Create(Value)
    End With
End Function
