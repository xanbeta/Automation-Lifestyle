Set oWSH = CreateObject("WScript.Shell")
 vbsInterpreter = "cscript.exe"

 Call ForceConsole()

 Function printf(txt)
    WScript.StdOut.WriteLine txt
 End Function

 Function printl(txt)
    WScript.StdOut.Write txt
 End Function

 Function scanf()
    scanf = LCase(WScript.StdIn.ReadLine)
 End Function

 Function wait(n)
    WScript.Sleep Int(n * 1000)
 End Function

 Function ForceConsole()
    If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
        oWSH.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
        WScript.Quit
    End If
 End Function

dateCode = Right("0" & Month(Date), 2) &  Right("0" & Day(Date), 2) &  Right(Year(Date), 2)
dailyLink = "CD" & dateCode & ".zip"


printf "hi. this is the datecode" & dailyLink

wait(4)
