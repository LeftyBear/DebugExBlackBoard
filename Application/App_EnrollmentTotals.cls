VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Presentation.DTOCollection")
Option Explicit
Private Type Member
    MainMales           As VBA.Collection
    MainFemales         As VBA.Collection
    MainClasses         As VBA.Collection
    MainGrades          As VBA.Collection
    SpecialMaleGrand    As Long
    SpecialFemaleGrand  As Long
    SpecialClasses      As VBA.Collection
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.MainMales = New VBA.Collection
    Set This.MainFemales = New VBA.Collection
    Set This.MainClasses = New VBA.Collection
    Set This.MainGrades = New VBA.Collection
    Set This.SpecialClasses = New VBA.Collection
End Sub

Public Sub AddMainMale(ByVal Grade As Long, ByVal Total As Long)
    Dim Key As String
    Key = CStr(Grade)
    Set This.MainMales.Item(Key) = Total
End Sub

Public Sub AddMainFemale(ByVal Grade As Long, ByVal Total As Long)
    Dim Key As String
    Key = CStr(Grade)
    Set This.MainFemales.Item(Key) = Total
End Sub

Public Sub AddMainClass(ByVal Grade As Long, ByVal ClassNo As Long, ByVal Total As Long)
    Dim Key As String
    Key = CStr(Grade) & CStr(ClassNo)
    Set This.MainClasses.Item(Key) = Total
    AddMainGrade Grade, Total
End Sub

Public Sub AddSpecialMale(ByVal Total As Long)
    This.SpecialMaleGrand = This.SpecialMaleGrand + Total
End Sub

Public Sub AddSpecialFemale(ByVal Total As Long)
    This.SpecialFemaleGrand = This.SpecialFemaleGrand + Total
End Sub

Public Sub AddSpecialClass(ByVal ClassNo As Long, ByVal Total As Long)
    Dim Key As String
    Key = CStr(ClassNo)
    Set This.SpecialClasses.Item(Key) = Total
End Sub

Public Function MainMaleTotalsCount() As Long
    MainMaleTotalsCount = This.MainMales.Count
End Function

Public Function MainFemaleTotalsCount() As Long
    MainFemaleTotalsCount = This.MainFemales.Count
End Function

Public Function MainClassTotalsCount() As Long
    MainClassTotalsCount = This.MainClasses.Count
End Function

Public Function MainGradeTotalsCount() As Long
    MainGradeTotalsCount = This.MainGrades.Count
End Function

Public Function SpecialClassTotalsCount() As Long
    SpecialClassTotalsCount = This.SpecialClasses.Count
End Function

Public Function GetMainMale(ByVal Grade As Long) As Long
    If This.MainMales.Count = 0 Then Exit Function
    Dim Key As String
    Key = CStr(Grade)
    If This.MainMales.Item(Key) Is Nothing Then Exit Function
    GetMainMale = This.MainMales.Item(Key)
End Function

Public Function GetMainFemale(ByVal Grade As Long) As Long
    If This.MainFemales.Count = 0 Then Exit Function
    Dim Key As String
    Key = CStr(Grade)
    If This.MainFemales.Item(Key) Is Nothing Then Exit Function
    GetMainFemale = This.MainFemales.Item(Key)
End Function

Public Function GetMainClass(ByVal Grade As Long, ByVal ClassNo As Long) As Long
    If This.MainClasses.Count = 0 Then Exit Function
    Dim Key As String
    Key = CStr(Grade) & CStr(ClassNo)
    If This.MainClasses.Item(Key) Is Nothing Then Exit Function
    GetMainClass = This.MainClasses.Item(Key)
End Function

Public Function GetMainGrade(ByVal Grade As Long) As Long
    If This.MainGrades.Count = 0 Then Exit Function
    Dim Key As String
    Key = CStr(Grade)
    If This.MainGrades.Item(Key) Is Nothing Then Exit Function
    GetMainGrade = This.MainGrades.Item(Key)
End Function

Public Function GetMainMaleGrand() As Long
    Dim i As Long
    For i = 1 To This.MainMales.Count
        Dim Total As Long
        Total = Total + This.MainMales.Item(i)
    Next
    GetMainMaleGrand = Total
End Function

Public Function GetMainFemaleGrand() As Long
    Dim i As Long
    For i = 1 To This.MainFemales.Count
        Dim Total As Long
        Total = Total + This.MainFemales.Item(i)
    Next
    GetMainFemaleGrand = Total
End Function

Public Function GetMainGradeGrand() As Long
    Dim i As Long
    For i = 1 To This.MainGrades.Count
        Dim Total As Long
        Total = Total + This.MainGrades.Item(i)
    Next
    GetMainGradeGrand = Total
End Function

Public Function GetSpecialMaleGrand() As Long
    GetSpecialMaleGrand = This.SpecialMaleGrand
End Function

Public Function GetSpecialFemaleGrand() As Long
    GetSpecialFemaleGrand = This.SpecialFemaleGrand
End Function

Public Function GetSpecialClass(ByVal ClassNo As Long) As Long
    If This.SpecialClasses.Count = 0 Then Exit Function
    Dim Key As String
    Key = CStr(ClassNo)
    If This.SpecialClasses.Item(Key) Is Nothing Then Exit Function
    GetSpecialClass = This.SpecialClasses.Item(Key)
End Function

Public Function GetSpecialGrand() As Long
    GetSpecialGrand = This.SpecialMaleGrand + This.SpecialFemaleGrand
End Function

Private Sub AddMainGrade(ByVal Grade As Long, ByVal Total As Long)
    Dim Key As String
    Key = CStr(Grade)
    Set This.MainGrades.Item(Key) = Total
End Sub
