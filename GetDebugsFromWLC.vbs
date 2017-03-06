' Bohack
' vWLC debug collector using send keys
' 2/28/17

If WScript.Arguments.Count = 0 Then
   Wscript.Echo "Usage: Script.vbs mac-address-list-file"
   WScript.Quit
End If

Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'total time =  minutes * intervals
'intervals = total time / minutes
 
minutes = 1
intervals = 42
index = 0
filename = WScript.Arguments(0)

Wscript.Echo "Every " & minutes & " minute(s) for " & intervals & " intervals with total of " & (intervals * minutes) & " minutes"
Wscript.Echo "30 seconds script will invoke config paging disable - Switch to vWLC console now!"
WScript.Sleep (30*1000)
WshShell.SendKeys "config paging disable"
WshShell.SendKeys "{enter}"
Do While index <= intervals
  WScript.Sleep (minutes*60*1000)
    WshShell.SendKeys "show time"
    WshShell.SendKeys "{enter}"
  Set f = fso.OpenTextFile(filename)
  Do Until f.AtEndOfStream
    WshShell.SendKeys "show client detail " & f.ReadLine
    WshShell.SendKeys "{enter}"
	WScript.Sleep 1000
  Loop
  f.Close
  Wscript.Echo "Interval " & (intervals - index) & " - " & time
  index = index + 1
Loop

Set fso = Nothing
Set WshShell = Nothing