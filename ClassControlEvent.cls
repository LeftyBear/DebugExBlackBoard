VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassControlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
Public Event OnClicked(ByVal ControlIndex As Long)
Public Event OnChanged(ByVal ControlIndex As Long)
Private ControlsCollection  As Collection
'@Ignore MoveFieldCloserToUsage
Private ControlsArray()     As ClassControlEventChild
Private Const TYPE_LABEL    As String = "Label"
Private Const TYPE_TEXTBOX  As String = "TextBox"
Private Const TYPE_COMBOBOX As String = "ComboBox"
Private Sub Class_Initialize()
    Set ControlsCollection = New Collection
End Sub
Public Sub AddControl(ByVal TargetControl As MSForms.Control)
    ControlsCollection.Add TargetControl
End Sub
Public Sub RegisterControls()
    If ControlsCollection.Count < 1 Then Exit Sub
    ReDim ControlsArray(1 To ControlsCollection.Count)
    Dim i As Long
    For i = 1 To ControlsCollection.Count
        Set ControlsArray(i) = New ClassControlEventChild
        With ControlsArray(i)
            Dim Target As MSForms.Control
            Set Target = ControlsCollection.Item(i)
            Select Case TypeName(Target)
                Case TYPE_LABEL
                    Set .LabelControl = MeLabel(Target.Name)
                Case TYPE_TEXTBOX
                    Set .TextBoxControl = MeTextBox(Target.Name)
                Case TYPE_COMBOBOX
                    Set .ComboBoxControl = MeComboBox(Target.Name)
            End Select
            .ItemIndex = i
            Set .RaiseEventClass = Me
        End With
    Next
End Sub
Public Function GetControlObject(ByVal ControlIndex As Long) As MSForms.Control
    Set GetControlObject = ControlsCollection.Item(ControlIndex)
End Function
Public Sub RaiseClick(ByVal ControlIndex As Long)
    RaiseEvent OnClicked(ControlIndex)
End Sub
Public Sub RaiseChange(ByVal ControlIndex As Long)
    RaiseEvent OnChanged(ControlIndex)
End Sub
Private Function MeLabel(ByVal ControlName As String) As MSForms.Label
    If Not ExistsControl(ControlName) Then Exit Function
    Set MeLabel = FormExBB.Controls.Item(ControlName)
End Function
Private Function MeTextBox(ByVal ControlName As String) As MSForms.TextBox
    If Not ExistsControl(ControlName) Then Exit Function
    Set MeTextBox = FormExBB.Controls.Item(ControlName)
End Function
Private Function MeComboBox(ByVal ControlName As String) As MSForms.ComboBox
    If Not ExistsControl(ControlName) Then Exit Function
    Set MeComboBox = FormExBB.Controls.Item(ControlName)
End Function
Private Function ExistsControl(ByVal ControlName As String) As Boolean
    Dim TargetControl As MSForms.Control
    For Each TargetControl In FormExBB.Controls
        If TargetControl.Name Like ControlName Then
            ExistsControl = True
            Exit Function
        End If
    Next
End Function

