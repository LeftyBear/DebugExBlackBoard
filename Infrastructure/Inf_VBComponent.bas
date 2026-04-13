Attribute VB_Name = "Inf_VBComponent"
'@Folder "Utility.VBComponent"
Option Explicit
Option Private Module

Private Const CtStdModule As Long = 1
Private Const CtClassModule As Long = 2
Private Const CtMsForm As Long = 3
Private Const RootPath As String = "C:\Users\biz\Documents\GitHub\DebugExBlackBoard"

Public Sub ExportAllModules()
    If App_Environment.GetEnvironmentType = ReleaseMode Then Exit Sub
    Dim Component As Object
    For Each Component In ThisWorkbook.VBProject.VBComponents
        If IsExportTarget(Component) Then ExportComponent RootPath, Component
    Next
    CleanupUnusedFiles RootPath
End Sub

Private Sub ExportComponent(ByVal RootPath As String, ByVal Component As Object)
    Dim FilePath As String
    FilePath = ResolveExportPath(RootPath, Component)
    If 0 < VBA.Len(VBA.Dir(FilePath)) Then VBA.Kill FilePath
    Component.Export FilePath
End Sub

Private Sub CleanupFolder(ByVal FolderPath As String)
    If VBA.Len(VBA.Dir(FolderPath, vbDirectory)) = 0 Then Exit Sub
    Dim FileName As String
    FileName = VBA.Dir(FolderPath & "*.*")
    Do While VBA.Len(FileName) > 0
        If FileName <> ".gitkeep" Then
            If Not IsModuleStillExists(RemoveExtension(FileName)) Then VBA.Kill FolderPath & FileName
        End If
        FileName = VBA.Dir
    Loop
End Sub

Private Sub CleanupUnusedFiles(ByVal RootPath As String)
    CleanupFolder RootPath & "\Domain\"
    CleanupFolder RootPath & "\Application\"
    CleanupFolder RootPath & "\Presentation\"
    CleanupFolder RootPath & "\Infrastructure\"
    CleanupFolder RootPath & "\Utility\"
End Sub

Private Function HasLayerPrefix(ByVal ModuleName As String) As Boolean
    If VBA.Left$(ModuleName, 4) = "Dom_" Then HasLayerPrefix = True
    If VBA.Left$(ModuleName, 4) = "App_" Then HasLayerPrefix = True
    If VBA.Left$(ModuleName, 4) = "Pre_" Then HasLayerPrefix = True
    If VBA.Left$(ModuleName, 4) = "Inf_" Then HasLayerPrefix = True
    If VBA.Left$(ModuleName, 5) = "Util_" Then HasLayerPrefix = True
End Function

Private Function IsExportTarget(ByVal Component As Object) As Boolean
    If Component.Type <> CtStdModule And Component.Type <> CtClassModule And Component.Type <> CtMsForm Then Exit Function
    If HasLayerPrefix(Component.Name) Then IsExportTarget = True
End Function

Private Function IsModuleStillExists(ByVal ModuleName As String) As Boolean
    Dim Component As Object
    For Each Component In ThisWorkbook.VBProject.VBComponents
        If Component.Name = ModuleName Then
            If IsExportTarget(Component) Then
                IsModuleStillExists = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function RemoveExtension(ByVal FileName As String) As String
    RemoveExtension = VBA.Left$(FileName, VBA.InStrRev(FileName, charPeriod) - 1)
End Function

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
        Err.Raise InfErrUnsupportedComponentType, "Util_VBComponent", "Unsupported component type."
    End Select
End Function

Private Function ResolveLayerFolder(ByVal RootPath As String, ByVal ModuleName As String) As String
    If ModuleName Like "Dom_*" Then
        ResolveLayerFolder = RootPath & "\Domain\"
    ElseIf ModuleName Like "App_*" Then
        ResolveLayerFolder = RootPath & "\Application\"
    ElseIf ModuleName Like "Pre_*" Then
        ResolveLayerFolder = RootPath & "\Presentation\"
    ElseIf ModuleName Like "Inf_*" Then
        ResolveLayerFolder = RootPath & "\Infrastructure\"
    ElseIf ModuleName Like "Util_*" Then
        ResolveLayerFolder = RootPath & "\Utility\"
    Else
        Err.Raise InfErrNotFoundLayerPrefix, "Util_VBComponent", "Layer prefix not found: " & ModuleName
    End If
End Function
