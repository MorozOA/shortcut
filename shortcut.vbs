' ================================================
' ==                                            ==
' ==               script version 1.00          ==
' ==                                            ==
' ==                                 Moroz Oleg ==
' ==                                 29/10/2012 ==
' ==                             mod 29/10/2012 ==
' ================================================

Dim oFS
set oFS = CreateObject("Scripting.FileSystemObject")

dim iDrvExists
iDrvExits = 0

if oFS.DriveExists("M:") Then
	iDrvExists = 1
End If

set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")

If (iDrvExist = 0) Then
	ret = WshShell.Run("cmd /c subst m: c:\", 0, TRUE)
End If
set oShellLink = WshShell.CreateShortcut(strDesktop & "\MyShortcut.lnk")
oShellLink.TargetPath = "M:\dir\start.bat"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll,43"
oShellLink.Description = "Shortcut Description text"
oShellLink.WorkingDirectory = "M:\dir\"
oShellLink.Save
If (iDrvExist = 0) Then
	ret = WshShell.Run("cmd /c subst m: /d", 0, TRUE)
End If