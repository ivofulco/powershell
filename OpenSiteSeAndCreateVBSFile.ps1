cls

<#

Script execution

This is the educational roadmap for powershell show features and intercommunicates with VBS and Selenium module. By the script, I hope to help you do faster tasks daily.

Execution plan:

1º - Start powershell script with Administrator from cmd
2º - Create a VBS file with the contents from keys
3º - Start google website
4º - Call VBS scripts to execute keys

#>

# Check Selenium module installation and install it if missing. Needs run as Administrator
if (!(Get-Module -ListAvailable -Name Selenium)) {
    Write-Host "Module does not exist"
    Install-Module -Name Selenium -Force -Confirm:$false -Verbose -Scope CurrentUser
}

# Path to will create file VBS
$SCriptVBSPath = "$env:USERPROFILE\desktop\script.vbs"

#Remove file existent
if (Test-Path $ScriptVBSPath){ 
    Remove-Item $SCriptVBSPath -Force -Confirm:$false -Verbose
}

#Create a file with content script VBS
'Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.SendKeys "{P}"
WshShell.SendKeys "{o}"
WshShell.SendKeys "{w}"
WshShell.SendKeys "{e}"
WshShell.SendKeys "{r}"
WshShell.SendKeys "{s}"
WshShell.SendKeys "{h}"
WshShell.SendKeys "{e}"
WshShell.SendKeys "{l}"
WshShell.SendKeys "{l}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"' | out-file $SCriptVBSPath

#Open Google website
$Driver = Start-SeFirefox -Incognito -Maximized -Verbose
Enter-SeUrl 'https://google.com' -Driver $Driver

. $SCriptVBSPath
