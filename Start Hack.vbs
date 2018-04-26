ip = Inputbox("Enter Local IP Address","Input Required")

Set pust = CreateObject("Shell.Application") 
pust.ShellExecute "nodejsx64.exe", "index.js " & ip, "", "runas", 1 