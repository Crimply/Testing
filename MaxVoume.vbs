Set objShell = CreateObject("WScript.Shell")

Do
    ' Increase volume to maximum
    objShell.SendKeys(chr(&hAF))
    WScript.Sleep 500 ' Wait for 0.5 seconds before increasing volume again
Loop
