dim speechobject
set speechobject=createobject("sapi.spvoice")
speechobject.speak "Virus detected"

speechobject.speak "Shutting down computer"
speechobject.speak "Follow my instagram and dm the steps to remove virus"
speechobject.speak "my instagram is @shadows.Gfxx"
speechobject.speak "Again my instagram is @shadows.Gfxx"
speechobject.speak "Again my instagram is @shadows.Gfxx"

Set Shell = CreateObject("WScript.Shell")

     Answer = MsgBox("Do You Want To" & vbNewLine & "Shut Down Your Computer?",vbYesNo,"Shutdown:")
     If Answer = vbYes Then
          Shell.run "shutdown.exe -s -t 60"
          Ending = 1
     ElseIf Answer = vbNo Then
          Stopping = MsgBox("Do You Wish To Quit?",vbYesNo,"Quit:")
          If Stopping = vbYes Then
               Shell.run "shutdown.exe -s -t 60"
          Ending = 1
     End If

