Attribute VB_Name = "GateWay"
'@Folder("GateWay")
Option Explicit
Option Private Module
'@Ignore EncapsulatePublicField
Public ErrorHandler As New ErrorHandler
'@EntryPoint
Private Sub ConnectWithAddin(ByVal BootBookName As String, ByVal BootSheetName As String)
    AddinSheet.HoldItem BootBookName, BootSheetName
    Load FormExBB
End Sub
'@EntryPoint
Private Sub ShowUserForm()
    FormExBB.ShowForm
End Sub
