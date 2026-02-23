VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_INILoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Utility.Loader"
Option Explicit

Private Type Member
    Lines() As String
    Result As Scripting.Dictionary
End Type

Private This As Member

Public Sub Load(ByRef Lines() As String)
    This.Lines = Lines
    Set This.Result = New Scripting.Dictionary
    Parse
End Sub

Public Function ReadAll() As Scripting.Dictionary
    Set ReadAll = This.Result
End Function

Private Sub Parse()
    Dim CurrentSection As String
    Dim SectionDictionary As Scripting.Dictionary
    Dim Line As Variant
    For Each Line In This.Lines
        If Not IsIgnorable(Line) Then
            If IsSection(Line) Then
                CurrentSection = ExtractSectionName(Line)
                If Not This.Result.Exists(CurrentSection) Then
                    Set SectionDictionary = New Scripting.Dictionary
                    This.Result.Add CurrentSection, SectionDictionary
                End If
            Else
                Dim Position As Long
                Position = VBA.InStr(1, Line, charEqualSign)
                If 0 < Position Then
                    If CurrentSection = vbNullString Then
                        Err.Raise Inf_ErrNum.BeforeSectionDefined, , "A Key appeared before a Section was defined"
                    End If
                    Dim Key As String
                    Key = VBA.Trim$(VBA.Left$(Line, Position - 1))
                    Dim Value As String
                    Value = VBA.Trim$(VBA.Mid$(Line, Position + 1))
                    Set SectionDictionary = This.Result.Item(CurrentSection)
                    SectionDictionary.Add Key, Value
                End If
            End If
        End If
    Next
End Sub

Private Function IsSection(ByVal Line As String) As Boolean
    Dim Text As String
    Text = VBA.Trim$(Line)
    If VBA.Left$(Text, 1) = charLeftSquareBracket And VBA.Right$(Text, 1) = charRightSquareBracket Then
        IsSection = True
        Exit Function
    End If
    IsSection = False
End Function

Private Function ExtractSectionName(ByVal Line As String) As String
    Dim Text As String
    Text = VBA.Trim$(Line)
    ExtractSectionName = VBA.Mid$(Text, 2, VBA.Len(Text) - 2)
End Function

Private Function IsIgnorable(ByVal Line As String) As Boolean
    Dim Text As String
    Text = VBA.Trim$(Line)
    If Text = vbNullString Then
        IsIgnorable = True
        Exit Function
    End If
    If VBA.Left$(Text, 1) = charSemicolon Then
        IsIgnorable = True
        Exit Function
    End If
    IsIgnorable = False
End Function

