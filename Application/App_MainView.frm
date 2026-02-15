VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} App_MainView 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "App_MainView.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "App_MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Application.View"
Option Explicit

'Private Type Member
'    SelectedDate As Date
'End Type

'Private This As Member

Private Sub SetGridValue(ByVal Kind As String, ByVal Value As Variant, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long)
    Dim Cell As Object
    Set Cell = ResolveGridControl(Kind, Grade, ClassNo)
    If Cell Is Nothing Then
        Exit Sub
    End If
    Cell.Text = CStr(Value)
End Sub

Private Function GetGridLongValue(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As Long
    Dim TextValue As String
    TextValue = GetTextFromGrid(Kind, Grade, ClassNo)
    If TextValue = vbNullString Then
        GetGridLongValue = 0
        Exit Function
    End If
    GetGridLongValue = CLng(TextValue)
End Function

Private Function GetGridStringValue(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As String
    GetGridStringValue = GetTextFromGrid(Kind, Grade, ClassNo)
End Function

Private Function GetTextFromGrid(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As String
    Dim Control As Object
    Set Control = ResolveGridControl(Kind, Grade, ClassNo)
    If Control Is Nothing Then
        GetTextFromGrid = vbNullString
        Exit Function
    End If
    GetTextFromGrid = CStr(Control.Text)
End Function

Private Function ResolveGridControl(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As Object
    Dim ControlName As String
    ControlName = BuildGridControlName(Kind, Grade, ClassNo)
    On Error Resume Next
    Set ResolveGridControl = Me.Controls(ControlName)
    On Error GoTo 0
End Function

Private Function BuildGridControlName(ByVal Kind As String, Optional ByVal Grade As Long, Optional ByVal ClassNo As Long) As String
    Dim Cells() As Variant
    Cells = Array(Kind, CStr(Grade), CStr(ClassNo))
    BuildGridControlName = VBA.Join(Cells, Util_Character.UnderScore)
End Function

