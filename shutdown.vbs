Set objShell = CreateObject("Wscript.Shell")
Set wshShell = WScript.CreateObject ("WSCript.shell")
Set wshNetwork = WScript.CreateObject("WScript.Network")
strShutdown = "shutdown -s -t 0 -f"

Continue = MsgBox ("Are you sure you want to shutdown?",vbYesNo + VBCritical, "Shutdown computer")
If Continue = vbNo then
	' Do nothing
Else
	' This box will disappear after 5 seconds.
	intReturn = objShell.Popup("The computer will now backup and shutdown.", 5, "Shutdown computer", VBInformation)
	wshshell.run "robocopy /r:0 /mir /z /s /xj C:\Users\" & WshNetwork.Username & "\Documents Z:\Documents", 6, True
	wshshell.run "robocopy /r:0 /mir /z /s /xj C:\Users\" & WshNetwork.Username & "\Pictures Z:\Pictures", 6, True

	' Shutdown the computer
	objShell.Run strShutdown
End If
