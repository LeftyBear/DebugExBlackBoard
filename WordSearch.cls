VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sheet")
Option Explicit
Implements IWordSearch
Private Function IWordSearch_FoundItemNameList(ByVal SearchWords As String, ByVal Sheet As ISheetWithTable) As Scripting.IDictionary
    Dim RET As Dictionary
    Set RET = New Dictionary
    Dim Wrods() As String
    Wrods = Split(Replace(SearchWords, WIDE_SPACE, HALF_SPACE))
    With Sheet
        Dim DataBody() As Variant
        DataBody = .AllRecords.Value
        Dim DateC As Long
        DateC = .HeaderItemColumns("“ú•t")
        Dim R As Long
        For R = LBound(DataBody, 1) To UBound(DataBody, 1)
            Dim C As Long
            For C = LBound(DataBody, 2) To UBound(DataBody, 2)
                Dim i As Long
                For i = LBound(Wrods) To UBound(Wrods)
                    If DataBody(R, C) Like "*" & Wrods(i) & "*" Then
                        Dim HitCount As Long
                        HitCount = HitCount + 1
                        If HitCount = UBound(Wrods) + 1 Then
                            RET.Add DataBody(R, DateC), .HeaderItemName(C)
                            Dim HasFound As Boolean
                            HasFound = True
                            Exit For
                        End If
                    End If
                Next
                HitCount = 0
                If HasFound Then Exit For
            Next
            HasFound = False
        Next
    End With
    Set IWordSearch_FoundItemNameList = RET
End Function
