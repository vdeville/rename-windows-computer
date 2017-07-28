' Author: Valentin DEVILLE
' Description: Rename computer with simply inputbox
' Licence: GPLV2

Title = "Renaming PC"

If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit
End If

Dim computername, newname
Set wshShell = CreateObject("WScript.Shell")
computername = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

newname = inputbox("What name you would like to apply ?", "Enter name", computername)

Set objWMIService = GetObject("Winmgmts:root\cimv2")

For Each objComputer in _
    objWMIService.InstancesOf("Win32_ComputerSystem")

        Return = objComputer.rename(newname)
        If Return <> 0 Then
           WScript.Echo "Rename failed. Error = " & Err.Number
        Else
           WScript.Echo "Rename succeeded." & _
               " Reboot for new name to go into effect"
        End If

Next


Question = MsgBox("PC will change name after restarting do you want ot restart now ?" & vbCrLf &_
"Yes to restart" & vbCrLF &_
"No to cancel" & vbCrLF, VbYesNo+VbQuestion, Title)
If Question = VbYes then 
    Set ws = CreateObject("Wscript.Shell")
	Command = "shutdown /r /t 1"
	Result = ws.run(Command, 0, True)
Else
    wscript.Quit(1)
End If
