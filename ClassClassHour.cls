VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassClassHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
Private ClassHoursList_ As Dictionary
Private TimeTables_ As Dictionary
Public Property Set ClassHoursList(ByVal Target As Dictionary)
    Set ClassHoursList_ = Target
End Property
Public Property Get ClassHoursList() As Dictionary
    Set ClassHoursList = ClassHoursList_
End Property
Public Property Set TimeTables(ByVal Target As Dictionary)
    Set TimeTables_ = Target
End Property
Public Property Get TimeTables() As Dictionary
    Set TimeTables = TimeTables_
End Property
Public Function TotalGrade(ByVal Grade As String) As Long
    Dim RET As Long
    Select Case Grade
        Case "1年", "2年", "3年"
            RET = Sum授業(Grade) + Sum行事(Grade) + Sum児童会(Grade)
        Case "4年"
            RET = Sum授業(Grade) + Sum行事(Grade) + Sum児童会(Grade) + Sumクラブ
        Case "5年", "6年"
            RET = Sum授業(Grade) + Sum行事(Grade) + Sum児童会(Grade) + Sumクラブ + Sum委員会
    End Select
    TotalGrade = RET
End Function
Private Function Sum授業(ByVal Grade As String) As Long
    Sum授業 = Me.ClassHoursList.Item("授業" & Grade)
End Function
Private Function Sum行事(ByVal Grade As String) As Long
    Sum行事 = Me.ClassHoursList.Item("行事" & Grade)
End Function
Private Function Sum児童会(ByVal Grade As String) As Long
    Sum児童会 = Me.ClassHoursList.Item("児童会" & Grade)
End Function
Private Function Sumクラブ() As Long
    Sumクラブ = Me.ClassHoursList.Item("クラブ")
End Function
Private Function Sum委員会() As Long
    Sum委員会 = Me.ClassHoursList.Item("委員会")
End Function
Public Function MismatchedClassHourMessage() As String
    Dim Messages As Collection
    Set Messages = New Collection
    Dim Grades() As Variant
    Grades = Array("1年", "2年", "3年", "4年", "5年", "6年")
    Dim i As Long
    For i = 0 To 5
        Dim MismatchedList As Collection
        Set MismatchedList = New Collection
        Select Case Grades(i)
            Case "1年", "2年", "3年"
                If Mismatched授業(Grades(i)) Then MismatchedList.Add "授業"
                If Mismatched行事(Grades(i)) Then MismatchedList.Add "行事"
                If Mismatched児童会(Grades(i)) Then MismatchedList.Add "児童会"
            Case "4年"
                If Mismatched授業(Grades(i)) Then MismatchedList.Add "授業"
                If Mismatched行事(Grades(i)) Then MismatchedList.Add "行事"
                If Mismatched児童会(Grades(i)) Then MismatchedList.Add "児童会"
                If Mismatchedクラブ(Grades(i)) Then MismatchedList.Add "クラブ"
            Case "5年", "6年"
                If Mismatched授業(Grades(i)) Then MismatchedList.Add "授業"
                If Mismatched行事(Grades(i)) Then MismatchedList.Add "行事"
                If Mismatched児童会(Grades(i)) Then MismatchedList.Add "児童会"
                If Mismatchedクラブ(Grades(i)) Then MismatchedList.Add "クラブ"
                If Mismatched委員会(Grades(i)) Then MismatchedList.Add "委員会"
        End Select
        If MismatchedList.Count > 0 Then
            Dim MismatchedArray() As Variant
            Dim Message As String
            MismatchedArray = LIB.CollectionItemsToArray(MismatchedList)
            Message = Grades(i) & "の " & Join(MismatchedArray, "、")
            Messages.Add Message
        End If
    Next
    If Messages.Count = 0 Then Exit Function
    MismatchedClassHourMessage = Join(LIB.CollectionItemsToArray(Messages), vbCrLf)
End Function
Private Function Count授業(ByVal Grade As String) As Long
    Dim RET As Long
    Dim Class As Variant
    For Each Class In TimeTables
        If Class Like "#校時" & Grade Then
            If Me.TimeTables.Item(Class) = "○" Then
                RET = RET + 1
            End If
        End If
    Next
    Count授業 = RET
End Function
Private Function Count行事(ByVal Grade As String) As Long
    Dim RET As Long
    Dim Class As Variant
    For Each Class In TimeTables
        If Class Like "#校時" & Grade Then
            If Me.TimeTables.Item(Class) = "行" Then
                RET = RET + 1
            End If
        End If
    Next
    Count行事 = RET
End Function
Private Function Count児童会(ByVal Grade As String) As Long
    Dim RET As Long
    Dim Class As Variant
    For Each Class In TimeTables
        If Class Like "#校時" & Grade Then
            If Me.TimeTables.Item(Class) = "児" Then
                RET = RET + 1
            End If
        End If
    Next
    Count児童会 = RET
End Function
Private Function Countクラブ(ByVal Grade As String) As Long
    Dim RET As Long
    Dim Class As Variant
    For Each Class In TimeTables
        If Class Like "#校時" & Grade Then
            If Me.TimeTables.Item(Class) = "ク" Then
                RET = RET + 1
            End If
        End If
    Next
    Countクラブ = RET
End Function
Private Function Count委員会(ByVal Grade As String) As Long
    Dim RET As Long
    Dim Class As Variant
    For Each Class In TimeTables
        If Class Like "#校時" & Grade Then
            If Me.TimeTables.Item(Class) = "委" Then
                RET = RET + 1
            End If
        End If
    Next
    Count委員会 = RET
End Function
Private Function Mismatched授業(ByVal Grade As String) As Boolean
    If Sum授業(Grade) = Count授業(Grade) Then Exit Function
    Mismatched授業 = True
End Function
Private Function Mismatched行事(ByVal Grade As String) As Boolean
    If Sum行事(Grade) = Count行事(Grade) Then Exit Function
    Mismatched行事 = True
End Function
Private Function Mismatched児童会(ByVal Grade As String) As Boolean
    If Sum児童会(Grade) = Count児童会(Grade) Then Exit Function
    Mismatched児童会 = True
End Function
Private Function Mismatchedクラブ(ByVal Grade As String) As Boolean
    If Sumクラブ = Countクラブ(Grade) Then Exit Function
    Mismatchedクラブ = True
End Function
Private Function Mismatched委員会(ByVal Grade As String) As Boolean
    If Sum委員会 = Count委員会(Grade) Then Exit Function
    Mismatched委員会 = True
End Function


