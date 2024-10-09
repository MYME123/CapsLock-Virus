Set wshShell = WScript.CreateObject("WScript.Shell")
Do
    WScript.Sleep 100
    wshShell.SendKeys "{CAPSLOCK}"
Loop
