Option Explicit
Dim Title,ws,nMinutes,nSeconds,sMessage,objShell, nPID
Set objShell = wscript.CreateObject("wscript.Shell")
Title = "Counting Down to Kill Process"
Set ws = CreateObject("wscript.Shell")
nPID = InputBox("Process PID to kill:","","PID")
if nPID = "PID" Then
msgbox("No PID number specified, please restart the script!")
Wscript.Quit
end if
if nPID = "" Then
msgbox("No PID number specified, please restart the script!")
Wscript.Quit
end if
nSeconds = InputBox("Seconds to kill process:","","60")
if nSeconds = "" Then
msgbox("No seconds specified, please restart the script!")
Wscript.Quit
end if
nMinutes = InputBox("Minutes to kill process:","","0")
if nMinutes = "" Then
msgbox("No minutes specified, please restart the script!")
Wscript.Quit
end if
sMessage = "<font color=Yellow style=font-family:consolas size=2><b>Counting Down to Kill Process " &nPID
'Open a chromeless window with message
with HTABox("Blue",100,300,0,630)
    .document.title = "Counting Down Notification to Kill Process"
    .msg.innerHTML = sMessage
    do until .done.value or (nMinutes + nSeconds < 1)
        .msg.innerHTML = sMessage & "<br>" & nMinutes & ":" & Right("0"&nSeconds, 2) _
        & " remaining</b></font><br>"
        wsh.sleep 1000 ' milliseconds
        nSeconds = nSeconds - 1
        if nSeconds < 0 then 
            if nMinutes > 0 then
                nMinutes = nMinutes - 1
                nSeconds = 59
            end if
        end if
    loop
    .done.value = true
    .close
end with
objShell.run("C:/Windows/system32/taskkill.exe /F /PID "&nPID)
'*****************************************************************
Function HTABox(sBgColor, h, w, l, t)
    Dim IE, HTA, sCmd, nRnd
    randomize : nRnd = Int(1000000 * rnd)
    sCmd = "mshta.exe ""javascript:{new " _
    & "ActiveXObject(""InternetExplorer.Application"")" _
    & ".PutProperty('" & nRnd & "',window);" _
    & "window.resizeTo(" & w & "," & h & ");" _
    & "window.moveTo(" & l & "," & t & ")}"""
    with CreateObject("WScript.Shell")
        .Run sCmd, 1, False
        do until .AppActivate("javascript:{new ") : WSH.sleep 10 : loop
        end with  'WSHShell
        For Each IE In CreateObject("Shell.Application").windows
            If IsObject(IE.GetProperty(nRnd)) Then
                set HTABox = IE.GetProperty(nRnd)
                IE.Quit
                HTABox.document.title = "HTABox"
                HTABox.document.write _
                "<HTA:Application contextMenu=no border=thin " _
                & "minimizebutton=no maximizebutton=no sysmenu=no SHOWINTASKBAR=no >" _
                & "<body scroll=no style='background-color:" _
                & sBgColor & ";font:normal 10pt Arial;" _
                & "border-Style:inset;border-Width:3px'" _
                & "onbeforeunload='vbscript:if not done.value then " _
                & "window.event.cancelBubble=true:" _
                & "window.event.returnValue=false:" _
                & "done.value=true:end if'>" _
                & "<input type=hidden id=done value=false>" _
                & "<center><span id=msg>&nbsp;</span><br>" _
                & "<input type=button id=btn1 value=' OK ' "_
                & "onclick=done.value=true><center></body>"
                HTABox.btn1.focus
                Exit Function
            End If
        Next
        MsgBox "HTA window not found."
        wsh.quit
End Function
'*****************************************************************
