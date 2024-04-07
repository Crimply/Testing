Set objShell = CreateObject("WScript.Shell")

Do
    ' Increase volume to maximum
    objShell.SendKeys(chr(&hAF))
    WScript.Sleep 50 ' Wait for 0.05 seconds before increasing volume again
Loop
