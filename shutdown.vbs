Option Explicit
Dim Title,ws,nMinutes,nSeconds,sMessage,objShell
Set objShell = wscript.CreateObject("wscript.Shell")
Title = "Counting Down to Shutdown"
Set ws = CreateObject("wscript.Shell")
nSeconds = InputBox("Seconds to shutdown:","","0")
nMinutes = InputBox("Minutes to shutdown:","","0")
'nMinutes = 0
'nSeconds = 0
sMessage = "<font color=Red size=2><b>Counting Down to Shutdown"
'Open a chromeless window with message
with HTABox("lightBlue",100,250,0,630)
    .document.title = "Counting Down Notification to Shutdown"
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
objShell.run("C:/Windows/system32/shutdown.exe /s /f")
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
