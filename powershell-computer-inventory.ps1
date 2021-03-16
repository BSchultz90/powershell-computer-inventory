
#Global Error Suppression
$ErrorActionPreference = "SilentlyContinue"

#Set the Date for Later Variables
$Date = Get-Date -Format "FileDate"

#This Code Archives the Previous Weeks Inventories | Then Checks and Removes Inventories Older than 30 days.
Get-ChildItem -Path '\\Server\Scripts\Full Site Inventories' | Move-Item -Destination '\\PHX00134\Support\Desktop\ITPT\Scripts\Full Site Inventories\Archive'
Get-ChildItem -Path '\\PHX00134\Support\Desktop\ITPT\Scripts\Full Site Inventories\Archive' | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-30)} | Remove-Item -Force

#The Below Code Initializes an Array of Lists from the Path Listed to out of PC Lists | It's then Iterated over in the For Loop to Push the Data into the Array.
$RootListPath = Get-ChildItem -Path '\\Server\PC_ALL' -Recurse | Select-Object -ExpandProperty 'Name'
$ArrayOfLists = @()

ForEach ($List in $RootListPath) {
    $ArrayOfLists += $list
}

$PSEmailServer = 'PHX01471.bhs.bannerhealth.com'

ForEach($List in $ArrayOfLists) {
   $CurrentList = Get-Content "Server\PC_ALL\$List"
   $CurrentSiteName = $List -replace ".txt",""
   $Results = ForEach($PC in $CurrentList) {
    If(Test-Connection -ComputerName $PC -Count 1 -Quiet ) {
    [PSCustomObject] @{
        PCName = Get-CimInstance -ComputerName $PC -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Name
        SerialNumber = Get-CimInstance -ComputerName $PC -ClassName Win32_BIOS | Select-Object -ExpandProperty SerialNumber 
        ModelName = Get-CimInstance -ComputerName $PC -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Model
        CRIT = Invoke-Command -ComputerName $PC {Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\" -Name "SrvComment" | Select-Object -ExpandProperty srvcomment}
        OS = Get-CimInstance -ComputerName $PC -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty Caption
        IPAddress = Test-Connection -ComputerName $PC -Count 1 | Select-Object -ExpandProperty IPV4Address | Select-Object -ExpandProperty IPAddresstoString
    } 
} } 
    $Results | Export-Csv -Path "\\PHX00134\Support\Desktop\ITPT\Scripts\Full Site Inventories\$CurrentSiteName Inventory $Date.csv" -NoTypeInformation
    Write-Host "`n Inventory for Site - $CurrentSiteName Has Completed."
}

#The Below Code Sends an Email Upon Completion #FIXME: Emails are Hardcoded >>> Change this to a Dedicated Email Address.
Send-MailMessage -To "Brice.Schultz@Email.com" -From "Brice.Schultz@Email.com" -Subject "Refresh Inventory Script has Completed" -Body "The Refresh Inventory Script for $location is Finished `n BEEP BOOP This Email was sent via PowerShell"

Exit
