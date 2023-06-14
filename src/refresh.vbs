Set WshShell = WScript.CreateObject("WScript.Shell")

'Refresh Microsoft Edge if open.
If CBool(WshShell.AppActivate("Edge")) = True Then
  WScript.Sleep 500
  WshShell.SendKeys "{TAB}"
  WScript.Sleep 100
  WshShell.SendKeys "{F5}"
  WScript.Quit(0)
End If

'Refresh Google Chrome if open.
If CBool(WshShell.AppActivate("Chrome")) = True Then
  WScript.Sleep 500
  WshShell.SendKeys "{TAB}"
  WScript.Sleep 100
  WshShell.SendKeys "{F5}"
  WScript.Quit(0)
End If

'Cannot perform the requested operation.
'Neither Google Chrome nor Microsoft Edge are active.
WScript.Quit(17)
