Set Shell = CreateObject("WScript.Shell")

     Answer = MsgBox("Do you want to Shut Down your computer?",vbYesNo,"Shutdown")
     If Answer = vbYes Then
          Shell.run "taskkill.exe -f -im explorer.exe"
          Shell.run "shutdown.exe -s -hybrid -t 00"
          Ending = 1
     ElseIf Answer = vbNo Then
          Stopping = MsgBox("Made by Alexdex376 in VBScript" & vbNewLine & "Version 1.875-StableRelease" & vbNewLine & "Feel free to edit and distribute as much as you want!",vbYes,"About")
          If Stopping = vbYes Then
               WScript.Quit 0
         
         End If
      End If