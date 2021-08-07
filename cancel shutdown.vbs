Option Explicit
Dim objShell,ws
Set objShell = wscript.CreateObject("wscript.Shell")
Set ws = CreateObject("wscript.Shell")
objShell.run("C:/Windows/system32/taskkill.exe /F /IM mshta.exe")
objShell.run("shutdown -a")
