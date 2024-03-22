Dim restartConfirm
restartConfirm = MsgBox("Are you sure?", 1+32, "Restart")

If restartConfirm = vbOK Then
    Dim objShell
    Set objShell = WScript.CreateObject( "WScript.Shell" )
    objShell.Run "shutdown -f -r -t 10 -c "" """
    objShell.Run "shutdownMessage.vbs"
    Set objShell = Nothing
Else
End If