Set WshShell = WScript.CreateObject("WScript.Shell")

'WshShell.Run "program"

message = InputBox("Message?")

num = InputBox("Enter the amount of messages")

sleep = InputBox("Gap between messages? Enter a number.")

msgbox("Now click into a text field and wait")

WScript.Sleep 2000 

Do

WshShell.SendKeys message
WshShell.SendKeys "{ENTER}"
WScript.Sleep sleep

num = num - 1

Loop Until num < 1

msgbox("fin")