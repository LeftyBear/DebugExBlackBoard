VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassHourItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sheet")
Option Explicit
Implements ICalculator
Private Enum ClassHourItemNameParts
    Grade = 1
End Enum
Private Enum TimeTableItemNameParts
    Period = 0
    Grade
End Enum
Implements IFinder
Private Const DELIMITER As String = "_"
Private Function ICalculator_Totaled(ByVal RecordList As Scripting.IDictionary) As Variant()
    Dim ItemName As Variant
    For Each ItemName In RecordList
        If ItemName Like "*合計*" Then
            Dim Grade As String
            Grade = Split(ItemName, DELIMITER)(ClassHourItemNameParts.Grade)
            RecordList.Item(ItemName) = Total学年(RecordList, Grade)
        End If
    Next
    Dim ListItems() As Variant
    ListItems = RecordList.Items
    ICalculator_Totaled = LIB.TransposeArray1D→2D(ListItems)
End Function
Private Function Total学年(ByVal RecordList As Dictionary, ByVal Grade As String) As Long
    Dim ItemName As Variant
    For Each ItemName In RecordList
        If ItemName Like "*" & Grade & "年" Then
            If Not ItemName Like "*合計*" Then
                Dim RET As Long
                RET = RET + RecordList.Item(ItemName)
            End If
        End If
    Next
    Total学年 = RET
End Function
Private Function IFinder_Find(ByVal RecordList As Scripting.IDictionary) As String
    Dim RET As Collection
    Set RET = New Collection
    Dim Grade As Long
    For Grade = 1 To 6
        Dim BUF As Collection
        Set BUF = MismatchedClassList(RecordList, Grade)
        If BUF.Count > 0 Then RET.Add Grade & "年の " & Join(LIB.CollectionItemsToArray(BUF), "、")
    Next
    If RET.Count = 0 Then Exit Function
    IFinder_Find = Join(LIB.CollectionItemsToArray(RET), vbCrLf)
End Function
Private Function MismatchedClassList(ByVal RecordList As Dictionary, ByVal Grade As Long) As Collection
    Dim RET As Collection
    Set RET = New Collection
    If Not Equals(RecordList, Grade, "授業", "○") Then RET.Add "授業"
    If Not Equals(RecordList, Grade, "行事", "行") Then RET.Add "行事"
    If Not Equals(RecordList, Grade, "児童会", "児") Then RET.Add "児童会"
    If 4 <= Grade Then
        If Not Equals(RecordList, Grade, "クラブ", "ク") Then RET.Add "クラブ"
        If 5 <= Grade Then
            If Not Equals(RecordList, Grade, "委員会", "委") Then RET.Add "委員会"
        End If
    End If
    Set MismatchedClassList = RET
End Function
Private Function Equals(ByVal RecordList As Dictionary, ByVal Grade As Long, ByVal ItemName As String, ByVal TimeTableName As String) As Boolean
    Dim FindItemName As String
    FindItemName = ItemName & DELIMITER & Grade & "年"
    Dim TargetClassHour As Long
    TargetClassHour = RecordList.Item(FindItemName)
    Dim TotalTimeTable As Long
    TotalTimeTable = CountedTimeTableName(RecordList, Grade, TimeTableName)
    If TargetClassHour = TotalTimeTable Then Equals = True
End Function
Private Function CountedTimeTableName(ByVal RecordList As Dictionary, ByVal Grade As Long, ByVal TimeTableName As String) As Long
    Dim RET As Long
    Dim Period As Long
    For Period = 1 To 6
        Dim FindItemNameParts(TimeTableItemNameParts.Grade) As String
        FindItemNameParts(TimeTableItemNameParts.Period) = Period & "校時"
        FindItemNameParts(TimeTableItemNameParts.Grade) = Grade & "年"
        Dim FindItemName As String
        FindItemName = Join(FindItemNameParts, DELIMITER)
        If RecordList.Item(FindItemName) Like TimeTableName Then RET = RET + 1
    Next
    CountedTimeTableName = RET
End Function

