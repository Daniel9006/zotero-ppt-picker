Option Explicit

Private Const PROJECT_ROOT As String = "C:\Users\daniel\OneDrive\Zotero_Add-In\Python\zotero-ppt-picker"
Private Const START_PICKER_CMD As String = PROJECT_ROOT & "\scripts\start_picker.cmd"

Public Sub LaunchZoteroPicker(control As IRibbonControl)
    RunZoteroPickerAction ""
End Sub

Public Sub ZoteroUpdateDocument(control As IRibbonControl)
    RunZoteroPickerAction "update-document"
End Sub

Public Sub ZoteroRewriteBibliography(control As IRibbonControl)
    RunZoteroPickerAction "rewrite-bibliography"
End Sub

Public Sub ZoteroSetBibliographyTarget(control As IRibbonControl)
    RunZoteroPickerAction "set-bibliography-target"
End Sub

Private Sub RunZoteroPickerAction(ByVal actionName As String)
    Dim shell As Object
    Dim q As String
    Dim cmd As String

    Set shell = CreateObject("WScript.Shell")

    ' Picker UI only:
    ' If the picker is already open, bring it to the foreground instead of
    ' starting a second picker instance.
    If Len(actionName) = 0 Then
        On Error Resume Next
        If shell.AppActivate("Zotero Picker") Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If

    q = Chr$(34)

    cmd = "cmd.exe /c " & q & q & START_PICKER_CMD & q

    If Len(actionName) > 0 Then
        cmd = cmd & " --action " & actionName
    End If

    cmd = cmd & q

    ' Window style 0 hides the transient command window.
    shell.Run cmd, 0, False
End Sub
