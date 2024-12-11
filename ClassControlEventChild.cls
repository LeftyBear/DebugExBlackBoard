VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassControlEventChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormExBB")
Option Explicit
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents TargetLabel      As MSForms.Label
Attribute TargetLabel.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents TargetTextBox    As MSForms.TextBox
Attribute TargetTextBox.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents TargetComboBox   As MSForms.ComboBox
Attribute TargetComboBox.VB_VarHelpID = -1
Private ItemIndex_          As Long
Private RaiseEventClass_    As ClassControlEvent
'@Ignore WriteOnlyProperty
Public Property Set LabelControl(ByVal Target As MSForms.Label)
    Set TargetLabel = Target
End Property
'@Ignore WriteOnlyProperty
Public Property Set TextBoxControl(ByVal Target As MSForms.TextBox)
    Set TargetTextBox = Target
End Property
'@Ignore WriteOnlyProperty
Public Property Set ComboBoxControl(ByVal Target As MSForms.ComboBox)
    Set TargetComboBox = Target
End Property
Public Property Let ItemIndex(ByVal IndexNumber As Long)
    ItemIndex_ = IndexNumber
End Property
Public Property Get ItemIndex() As Long
    ItemIndex = ItemIndex_
End Property
Public Property Set RaiseEventClass(ByVal Target As ClassControlEvent)
    Set RaiseEventClass_ = Target
End Property
Public Property Get RaiseEventClass() As ClassControlEvent
    Set RaiseEventClass = RaiseEventClass_
End Property
Private Sub TargetLabel_Click()
    Me.RaiseEventClass.RaiseClick Me.ItemIndex
End Sub
Private Sub TargetLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.RaiseEventClass.RaiseClick Me.ItemIndex
End Sub
Private Sub TargetTextBox_Change()
    Me.RaiseEventClass.RaiseChange Me.ItemIndex
End Sub
Private Sub TargetComboBox_Change()
    Me.RaiseEventClass.RaiseChange Me.ItemIndex
End Sub

