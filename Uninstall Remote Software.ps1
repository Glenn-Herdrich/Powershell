[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

# Obtain hostname from user inputbox
DO {
    $title = 'Hostname'
    $msg   = 'Enter the FQDN or IP of host you wish to view/uninstall software on:'
    $hosts = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
} While ([string]::IsNullOrEmpty($hosts))

# View installed software or uninstall software
$msgBoxInput = [Microsoft.VisualBasic.Interaction]::MsgBox("Do you wish to look at installed software first",'YesNo,Question', "View Installed Software?")
    switch ($msgBoxInput) {
        'Yes' {
            Write-Host "Looking up software on " $hosts " please wait" 
            $apps = Get-WmiObject -Namespace "root\cimv2" -Class Win32_Product -ComputerName $hosts
            $apps | Out-GridView
            }
    }
# Obtain software name from user input
DO {
    $title = 'Software'
    $msg   = 'Enter the Win32_Product software name:'
    $software = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
} While ([string]::IsNullOrEmpty($software))

Write-Host "Processing search for " $software " on system " $hosts
$apps = Get-WmiObject -Namespace "root\cimv2" -Class Win32_Product -ComputerName $hosts | Where-Object {$_.Name -like '*' + $software + '*'}

If ([string]::IsNullOrEmpty($apps)) {
    [System.Windows.Forms.MessageBox]::Show("Match not found. Please check if entry matches Win32_Product name", "Not Found")
} Else {
    $msgBoxInput = [System.Windows.MessageBox]::Show('Do you wish to uninstall the listed software? ' + $apps,'Uninstall','YesNo','Error')
    switch ($msgBoxInput) {
        'Yes' {$apps.Uninstall()}
        'No' {exit}
    }
    $apps = Get-WmiObject -Namespace "root\cimv2" -Class Win32_Product -ComputerName $hosts | Where-Object {$_.Name -like '*' + $software + '*'}
Write-Host $apps
    If ([string]::IsNullOrEmpty($apps)) {
        [System.Windows.Forms.MessageBox]::Show("Software uninstalled. Exiting", "Uninstalled")
    }
}
