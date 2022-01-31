#################################### Sophos_V2_Upgrade ############################
##### Created by: Make-A-Wish Canada                                              #
##### Purpose:To upgrade the current LCWF VPN client to the new Sophos Connect V2 #
###################################################################################

$DATE = Get-Date -Format "dddd MM/dd/yyyy" 
$BACKUP = "$ENV:ProgramData\VPN_BACKUP"
$LOGFILE = "$ENV:ProgramData\VPN_BACKUP\LOG.txt"
$OLD_VPNBASE_PATH = "${env:ProgramFiles(x86)}\Sophos\Sophos SSL VPN Client"
$OLD_CONFIG_FILE = "${env:ProgramFiles(x86)}\Sophos\Sophos SSL VPN Client\config\*.ovpn"
$OLD_UNINSTALL_FILE = "${env:ProgramFiles(x86)}\Sophos\Sophos SSL VPN Client\Uninstall.exe"
$SOPHOS_V2_URL = "https://raw.githubusercontent.com/Cowboy1543/Test/main/SCV2.msi" ################## SOPHOS CONNECT V2 DOWNLOAD LINK 
$SOPHOS_SCCLI = "${env:ProgramFiles(x86)}\Sophos\Connect\sccli.exe"
$SOPHOS_CONNECTV2 = "${env:ProgramFiles(x86)}\Sophos\Connect\GUI\scgui.exe"
$SETUP_FILE = $env:HOMEDRIVE+$env:HOMEPATH+"\SCV2.msi"

Try {
    if ((Test-Path -Path $OLD_CONFIG_FILE) -eq $true){
        $output += "Running script Sophos_V2_Upgrade : "+$DATE 
        ############kill all current running sophos services and programs############
        $services = Get-Service -Name "OpenVPN*"
        $process = Get-Process "openvpn-gui"
        $server = Get-Process "openvpnserv"
        if ($process) {
            Stop-Process $process -force
            $output += "`nStatus: Sophos process closed"
        }
        if ($server) {
            Stop-Process $server -force 
            $output += "`nStatus: Sophos server process closed"
        }
        if ($services.Status -eq "Running"){
            if ($services.count -eq 1){
                stop-service -name $services.Name  -Force 

                $output += "`nStatus: Services closed "
            }else {
                $services | ForEach-Object{
                    Stop-Service -Name $_.Name -Force 
                }
                $output += "`nStatus: Services closed "
            }
        }
        ############backup old vpn config file############
        If ((Test-Path -Path $BACKUP) -eq $true){
            #copy old config file to this dir incase of any issues 
            Copy-Item -Path $OLD_CONFIG_FILE -Destination $BACKUP -Force
            $output += "`nStatus: Config file backed up "
        }else{
            #create backup DIR and copy old config file 
            #making it a variable to remove output 
            $null = New-Item -Path "$ENV:ProgramData " -Name "VPN_BACKUP" -ItemType "directory" -Force
            Copy-Item -Path $OLD_CONFIG_FILE -Destination $BACKUP -Force
            $output += "`nStatus: Config file backed up "
        }
        ############uninstall old client & DIR clean up############
        if ((Test-Path -Path $OLD_UNINSTALL_FILE) -eq $true){
            Start-Process $OLD_UNINSTALL_FILE -ArgumentList /S -Wait
            $output += "`nStatus: Old client uninstalled "
            Remove-Item -Path $OLD_VPNBASE_PATH -Recurse -Force
            $output += "`nStatus: Removing old directory "
        } else{
            $output += "`nWarning: SSL vpn uninstall file not found. The program might be uninstalled already "
            if ((Test-Path -Path $OLD_VPNBASE_PATH) -eq $true){ 
                Remove-Item -Path $OLD_VPNBASE_PATH -Recurse -Force
                $output += "`nStatus: Removing old directory "
            }
        }
        ############Download new client############ 
        try{
            Invoke-WebRequest -Uri $SOPHOS_V2_URL -OutFile $SETUP_FILE 
        }catch{$output += "`nERROR: Could not complete WebRequest , "+$_.ScriptStackTrace+"  MESSAGE: "+$_ + $DATE }
        ############Install new client############
        if ((Test-Path -path $SETUP_FILE) -eq $True){
            try{
            Start-Process $SETUP_FILE -ArgumentList /QN -Wait
            $output += "`nStatus: New client installed "
            }catch{$output += "`nERROR: ERROR WITH NEW CLIENT INSTALL, "+$_.ScriptStackTrace+"  MESSAGE: "+$_ + $DATE}
            #Remove Sophos connect V2 setup file
            Remove-Item -Path $SETUP_FILE -Force
            $output += "`nStatus: Setup file removed"
        }
        ############Import previous connection into new client############
        $file = get-item $BACKUP"\*.ovpn"
        if ((Test-Path -path $SOPHOS_SCCLI) -eq $True){
            try{
            $import_config = CMD /c "`"$SOPHOS_SCCLI`" add -f "$file
            $output += "`nStatus: Old config file imported"
            }catch{$output += "`nERROR: Connection was not imported"}
        }else {
            $output += "`nERROR: SCCLI file not found. Connection import not possible"
        }
        ############Start new client############
        try{
            Start-Process $SOPHOS_CONNECTV2
        }catch{ $output += "`nERROR: Could not start the new client"}
        ############ Write to log file############
        $output += "`nStatus: End of script.`n"
        $output | Out-File -FilePath $LOGFILE -Append
}else {
    #No old ssl vpn file found. 
    If ((Test-Path -Path $BACKUP) -ne $true){
        $null = New-Item -Path "$ENV:ProgramData " -Name "VPN_BACKUP" -ItemType "directory" -Force
        $output += "`nERROR: Old config file not found. User might not have the Sophos SSL VPN installed." 
    }else{
        $output += "`nERROR: Old config file not found. User might not have the Sophos SSL VPN installed." 
    }
    $output | Out-File -FilePath $LOGFILE -Append
    Remove-Variable output
}
}Catch {$output = "`nERROR: ERROR WITH MAIN OPERATION, "+$_.ScriptStackTrace+"  MESSAGE: "+$_ + $DATE | Out-File -FilePath $LOGFILE -Append }