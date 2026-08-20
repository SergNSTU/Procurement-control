Option Explicit

Dim shell, fso, baseDir, launcherPath, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
launcherPath = fso.BuildPath(baseDir, "app\Start-Portable.ps1")

If Not fso.FileExists(launcherPath) Then
    MsgBox "Не найдена папка app." & vbCrLf & _
           "Скачайте и распакуйте репозиторий целиком, затем повторите запуск.", _
           vbCritical, "Procurement Control"
    WScript.Quit 1
End If

shell.CurrentDirectory = baseDir
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File " & Chr(34) & launcherPath & Chr(34)
shell.Run command, 0, False
