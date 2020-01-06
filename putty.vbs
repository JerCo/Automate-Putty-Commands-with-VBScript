Dim Password
Dim Destination

UserName = InputBox("Please Enter Your UserID:")
Password  = InputBox("Please Enter Your Password:")
Destination = InputBox("Please Enter Your Hostname or IP address")

Set Shell = WScript.CreateObject("WScript.Shell")
loginscript = "C:\Program Files\PuTTY\putty.exe -ssh " & UserName & "@" & Destination & " -pw " & Password

Shell.Exec(loginscript) 
WScript.Sleep 5000
Shell.AppActivate "HOSTNAME - PuTTY"
Shell.SendKeys "whoami"
Shell.SendKeys "{ENTER}"
