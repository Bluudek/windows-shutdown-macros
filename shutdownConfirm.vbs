Dim shutdownConfirm
shutdownConfirm = MsgBox("Are you sure?", 1+32, "Shutdown")

If shutdownConfirm = vbOK Then
    Dim objShell
    Set objShell = WScript.CreateObject( "WScript.Shell" )
    objShell.Run "shutdown -f -s -t 10 -c "" """
    objShell.Run "shutdownMessage.vbs"
    Set objShell = Nothing
Else
End If