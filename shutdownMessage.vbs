Dim abort
abort = MsgBox("Shutting down in 10 seconds. Abort?", 4+48, "Shutting down.")

If abort = vbYes Then
    Dim objShell
    Set objShell = WScript.CreateObject( "WScript.Shell" )
    objShell.Run "shutdown -a"
    Set objShell = Nothing
Else
End If
