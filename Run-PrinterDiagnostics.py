import os
import subprocess
from time import sleep
import win32com.shell.shell as shell
import wmi
import base64

####### Variables ##########
activeUser = os.getlogin() 
curUser = os.environ.get('USERNAME')
printers = subprocess.getoutput('C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -command "Get-Printer | Where name -notlike *microsoft* | where name -notlike *OneNote* | where name -notlike *fax* | select name,portname,drivername"')
printLogStatus = subprocess.getoutput('C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -command (Get-WinEvent -Listlog "Microsoft-Windows-PrintService/Operational).isEnabled"')
############################

###### Functions ##########
def prCyan(skk): print("\033[96m {}\033[00m" .format(skk))
def prGreen(skk): print("\033[92m {}\033[00m" .format(skk))
def prRed(skk): print("\033[91m {}\033[00m" .format(skk))
def prPurple(skk): print("\033[95m {}\033[00m" .format(skk))
def Welcome():
    print("####################################################################################################")
    prCyan("Run-PrintDiagnostics Version 1.1, Now in Python!\n Author: David Just")
    prGreen(f"Welcome to the Printer Diagnostics Utility \n Where we try to make printers slightly less painful! \n Currently Running as {curUser.upper()}")
    print("#################################################################################################### \r")
    if activeUser != curUser:
        prRed("Caution, you are currently running as curUser./nYou must run this tool as the logged on user in order to work with shared printers \nRunning as SYSTEM or another user will only show system wide printers!")
    sleep(1)

def FetchPrintLogs():       
     if printLogStatus != 'True':
         prGreen("System Log [Microsoft-Windows-PrintService/Operational] is disabled...Trying to enable now")
         try :
                shell.ShellExecuteEx(lpVerb = 'runas', lpFile = 'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe', lpParameters="$log = Get-WinEvent -ListLog Microsoft-Windows-PrintService/Operational;$log.IsEnabled=$true;$log.SaveChanges()")
                prGreen("Successfully Enabled Logging!\nPlease try to print a test page and check back for event logs.\n")
                sleep(1)
         except : 
                print("Failed to enable log")
                sleep(1)
     else:
        os.system("powershell.exe -command \"Get-WinEvent -LogName Microsoft-Windows-PrintService/Operational | Out-GridView ; pause\"")

def RestartSpooler():
    print("Restarting Print Spooler...")
    shell.ShellExecuteEx(lpVerb = 'runas', lpFile = 'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe', lpParameters="Get-Service Spooler | Stop-Service -force ; cmd /c \"del %systemroot%\System32\spool\printers* /Q\" ; Get-Service Spooler | Start-Service")

def PrintTestPage():
     wmi_o = wmi.WMI('.')
     wql = ("SELECT * FROM Win32_Printer")
     wmiPrinters = wmi_o.query(wql)
     i = 1
     for a in wmiPrinters:
         print(i, a.DeviceID)
         i+=1
     selection = int(input("Select a printer to send a test page to:"))   
     index = int(selection - 1)
     selectedPrinter = str(wmiPrinters[index].DeviceID)
     command = "Invoke-CimMethod -MethodName PrintTestPage -InputObject (Get-CimInstance win32_printer | where name -like " + '"' + selectedPrinter + '"' + ')'
     encodedBytes = command.encode("utf-16LE")
     encodedStr = base64.b64encode(encodedBytes)
     subprocess.call("C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -encodedcommand " + encodedStr.decode())
     
def OpenPrinterCP(Option):
	switch = {
        'Modern': os.system("cmd.exe /c \"start ms-settings:printers\""),
        'Classic': os.system("cmd.exe /c \"start  control printers\"")
    }
def AddNewPrinter():
    os.system("cmd.exe /c \"C:\\Windows\\System32\\rundll32.exe PRINTUI.DLL, PrintUIEntry /im\"")

def EndProgram():
    print("Thank you for using the python Print Diagnostic Tool. Happy Printing!")
    sleep(1)
    exit(0)

def MainMenu():
    if printLogStatus == 'True':
        Status = 'Enabled'
    else:
        Status = 'Disabled'
    prCyan(f"Printer Service Log Status: {Status}")
    print ("Installed Printers:")
    print (printers)
    print("*Main Menu*")
    prCyan("\rOptions: \n[1] Review Print Service Logs\n[2] Restart and Clear Print Spooler\n[3] Add a new printer\n[4] Open Classic Printer Control Panel\n[5] Open Modern Printer Settings Page\n[6] Print Test Page\n[7] Quit") 
    try :
        MenuSelection = int(input("Please Enter A Selection:"))
        match MenuSelection:            
            case 1: 
                FetchPrintLogs()
            case 2: 
                RestartSpooler()
            case 3: 
                AddNewPrinter()
            case 4: 
                OpenPrinterCP("Classic")
            case 5: 
                OpenPrinterCP("Modern")
            case 6: 
                PrintTestPage()
            case 7: 
                EndProgram()
    except ValueError:
        prRed("Please enter a valid selection")
        os.system("pause")
        MainMenu()
    
        
Welcome()    
while "True":
    MainMenu()
