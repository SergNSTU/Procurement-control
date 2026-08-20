Option Explicit

Dim shell, fso, baseDir, launcherPath, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
launcherPath = fso.BuildPath(baseDir, "Procurement_control_portable\Start.vbs")

If Not fso.FileExists(launcherPath) Then
    MsgBox "Не найдена папка Procurement_control_portable." & vbCrLf & _
           "Скачайте и распакуйте репозиторий целиком, затем повторите запуск.", _
           vbCritical, "Procurement Control"
    WScript.Quit 1
End If

shell.CurrentDirectory = fso.BuildPath(baseDir, "Procurement_control_portable")
command = "wscript.exe " & Chr(34) & launcherPath & Chr(34)
shell.Run command, 0, False
