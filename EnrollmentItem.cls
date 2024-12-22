VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnrollmentItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sheet")
Option Explicit
Implements ICalculator
Private Enum ItemNameParts
    Grade = 0
    Class
    Sex
End Enum
Private Const DELIMITER As String = "_"
Private Function ICalculator_Totaled(ByVal RecordList As Scripting.IDictionary) As Variant()
    Dim ItemName As Variant
    For Each ItemName In RecordList
        If ItemName Like "*合計" Then
            Dim Grade As String
            Grade = Split(ItemName, DELIMITER)(ItemNameParts.Grade)
            If Not Grade Like "全学年" Then
                Dim JudgmentString As String
                JudgmentString = Split(ItemName, DELIMITER)(ItemNameParts.Grade + 1)
                Select Case True
                    Case JudgmentString Like "#組"
                        RecordList.Item(ItemName) = Total学年の学級(RecordList, Grade, JudgmentString)
                    Case JudgmentString Like "[男女]"
                        RecordList.Item(ItemName) = Total学年の性別(RecordList, Grade, JudgmentString)
                    Case JudgmentString Like "男女"
                        RecordList.Item(ItemName) = Total学年(RecordList, Grade)
                End Select
            Else
                JudgmentString = Split(ItemName, DELIMITER)(ItemNameParts.Grade + 1)
                Select Case True
                    Case JudgmentString Like "[男女]"
                        RecordList.Item(ItemName) = Total全学年の性別(RecordList, JudgmentString)
                    Case JudgmentString Like "男女"
                        RecordList.Item(ItemName) = Total全学年(RecordList)
                End Select
            End If
        End If
    Next
    Dim ListItems() As Variant
    ListItems = RecordList.Items
    ICalculator_Totaled = LIB.TransposeArray1D→2D(ListItems)
End Function
Private Function Total学年の学級(ByVal RecordList As Dictionary, ByVal Grade As String, ByVal Class As String) As String
    Dim FindItemNameParts(ItemNameParts.Sex) As String
    FindItemNameParts(ItemNameParts.Grade) = Grade
    FindItemNameParts(ItemNameParts.Class) = Class
    FindItemNameParts(ItemNameParts.Sex) = "[男女]"
    Dim FindItemName As String
    FindItemName = Join(FindItemNameParts, DELIMITER)
    Total学年の学級 = SumUpedFindItems(RecordList, FindItemName)
End Function
Private Function Total学年の性別(ByVal RecordList As Dictionary, ByVal Grade As String, ByVal Sex As String) As String
    Dim FindItemNameParts(ItemNameParts.Sex) As String
    FindItemNameParts(ItemNameParts.Grade) = Grade
    FindItemNameParts(ItemNameParts.Class) = "#組"
    FindItemNameParts(ItemNameParts.Sex) = Sex
    Dim FindItemName As String
    FindItemName = Join(FindItemNameParts, DELIMITER)
    Total学年の性別 = SumUpedFindItems(RecordList, FindItemName)
End Function
Private Function Total学年(ByVal RecordList As Dictionary, ByVal Grade As String) As String
    Dim FindItemNameParts(ItemNameParts.Sex) As String
    FindItemNameParts(ItemNameParts.Grade) = Grade
    FindItemNameParts(ItemNameParts.Class) = "#組"
    FindItemNameParts(ItemNameParts.Sex) = "[男女]"
    Dim FindItemName As String
    FindItemName = Join(FindItemNameParts, DELIMITER)
    Total学年 = SumUpedFindItems(RecordList, FindItemName)
End Function
Private Function Total全学年の性別(ByVal RecordList As Dictionary, ByVal Sex As String) As String
    Dim FindItemNameParts(ItemNameParts.Sex) As String
    FindItemNameParts(ItemNameParts.Grade) = "??"
    FindItemNameParts(ItemNameParts.Class) = "#組"
    FindItemNameParts(ItemNameParts.Sex) = Sex
    Dim FindItemName As String
    FindItemName = Join(FindItemNameParts, DELIMITER)
    Total全学年の性別 = SumUpedFindItems(RecordList, FindItemName)
End Function
Private Function Total全学年(ByVal RecordList As Dictionary) As String
    Dim FindItemNameParts(ItemNameParts.Sex) As String
    FindItemNameParts(ItemNameParts.Grade) = "??"
    FindItemNameParts(ItemNameParts.Class) = "#組"
    FindItemNameParts(ItemNameParts.Sex) = "[男女]"
    Dim FindItemName As String
    FindItemName = Join(FindItemNameParts, DELIMITER)
    Total全学年 = SumUpedFindItems(RecordList, FindItemName)
End Function
Private Function SumUpedFindItems(ByVal RecordList As Dictionary, ByVal FindItemName As String) As String
    Dim ItemName As Variant
    For Each ItemName In RecordList
        If ItemName Like FindItemName Then
            If RecordList.Item(ItemName) Like vbNullString Then Exit Function
            Dim BUF As Long
            BUF = BUF + RecordList.Item(ItemName)
        End If
    Next
    SumUpedFindItems = CStr(BUF)
End Function

