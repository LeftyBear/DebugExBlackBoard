Attribute VB_Name = "Util_VBComponent"
'@Folder "Utility.VBComponent"
Option Explicit

Private Const CtStdModule As Long = 1
Private Const CtClassModule As Long = 2
Private Const CtMsForm As Long = 3
Private Const RootPath As String = "C:\Users\biz\Documents\GitHub\DebugExBlackBoard"

Public Sub ExportAllModules()
    If Util_Environment.GetEnvironment = ReleaseMode Then Exit Sub
    Dim Component As Object
    For Each Component In ThisWorkbook.VBProject.VBComponents
        If IsExportTarget(Component) Then ExportComponent RootPath, Component
    Next
End Sub

Private Function IsExportTarget(ByVal Component As Object) As Boolean
    If Component.Type <> CtStdModule And Component.Type <> CtClassModule And Component.Type <> CtMsForm Then Exit Function
    If HasLayerPrefix(Component.Name) Then IsExportTarget = True
End Function

Private Function HasLayerPrefix(ByVal ModuleName As String) As Boolean
    If VBA.Left(ModuleName, 4) = "Dom_" Then HasLayerPrefix = True
    If VBA.Left(ModuleName, 4) = "App_" Then HasLayerPrefix = True
    If VBA.Left(ModuleName, 4) = "Inf_" Then HasLayerPrefix = True
    If VBA.Left(ModuleName, 5) = "Util_" Then HasLayerPrefix = True
End Function

Private Sub ExportComponent(ByVal RootPath As String, ByVal Component As Object)
    Dim FilePath As String
    FilePath = ResolveExportPath(RootPath, Component)
    If 0 < VBA.Len(VBA.Dir(FilePath)) Then VBA.Kill FilePath
    Component.Export FilePath
End Sub

Private Function ResolveExportPath(ByVal RootPath As String, ByVal Component As Object) As String
    Dim LayerFolder As String
    LayerFolder = ResolveLayerFolder(RootPath, Component.Name)
    Select Case Component.Type
    Case CtStdModule
        ResolveExportPath = LayerFolder & Component.Name & ".bas"
    Case CtClassModule
        ResolveExportPath = LayerFolder & Component.Name & ".cls"
    Case CtMsForm
        ResolveExportPath = LayerFolder & Component.Name & ".frm"
    Case Else
        Err.Raise vbObjectError + 1000, "Util_ExportAllModules", "Unsupported component type."
    End Select
End Function

Private Function ResolveLayerFolder(ByVal RootPath As String, ByVal ModuleName As String) As String
    Select Case True
    Case VBA.Left$(ModuleName, 4) = "Dom_"
        ResolveLayerFolder = RootPath & "\Domain\"
    Case VBA.Left$(ModuleName, 4) = "App_"
        ResolveLayerFolder = RootPath & "\Application\"
    Case VBA.Left$(ModuleName, 4) = "Inf_"
        ResolveLayerFolder = RootPath & "\Infrastructure\"
    Case VBA.Left$(ModuleName, 5) = "Util_"
        ResolveLayerFolder = RootPath & "\Utility\"
    Case Else
        Err.Raise vbObjectError + 1001, "ResolveLayerFolder", "Layer prefix not found: " & ModuleName
    End Select
End Function
