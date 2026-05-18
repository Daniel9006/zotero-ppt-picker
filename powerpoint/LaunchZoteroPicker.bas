Attribute VB_Name = "LaunchZoteroPicker"
Option Explicit

' PowerPoint VBA starter for zotero-ppt-picker.
' Maintainer-facing comments are English; user-facing messages are German.
'
' Installation:
' 1. Import this .bas file into the PowerPoint VBA editor.
' 2. Adjust PICKER_LAUNCHER_PATH to your local repository path.
' 3. Run LaunchZoteroPicker from PowerPoint, or assign it to a button.

Private Const PICKER_LAUNCHER_PATH As String = "C:\Path\To\zotero-ppt-picker\scripts\start_picker.cmd"

Public Sub LaunchZoteroPicker()
    Dim launcherPath As String
    launcherPath = PICKER_LAUNCHER_PATH

    If Len(Dir$(launcherPath, vbNormal)) = 0 Then
        MsgBox "Der Zotero-Picker-Launcher wurde nicht gefunden:" & vbCrLf & _
               launcherPath & vbCrLf & vbCrLf & _
               "Bitte passe PICKER_LAUNCHER_PATH im VBA-Modul an.", _
               vbExclamation, _
               "Zotero Picker starten"
        Exit Sub
    End If

    On Error GoTo LaunchFailed
    Shell "cmd.exe /c """ & launcherPath & """", vbNormalFocus
    Exit Sub

LaunchFailed:
    MsgBox "Der Zotero-Picker konnte nicht gestartet werden." & vbCrLf & _
           "Launcher:" & vbCrLf & launcherPath & vbCrLf & vbCrLf & _
           "Fehler: " & Err.Description, _
           vbExclamation, _
           "Zotero Picker starten"
End Sub
