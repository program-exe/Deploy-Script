<#
.SYNOPSIS
    Ability to create a Virtual Machine (VM) from scratch with information from ServiceNow
.DESCRIPTION
    Performs the following action in order to build the VM:
        Gathers information from a CSV regarding the VM that needs to be created. Determines tags for VM that are needed.
        Clones and creates a template for the VM in the targeted vCenter
        Edits configuration of VM including: Name/description, numa information, networking information, adding tags and running OS Config script
.NOTES
  Version:        2.0
  Author:         <Joshua Dooling, Jason Nagin, Sean McConnell, Andrew Hosler, Matthew Johnson>
  Creation Date:  <07/16/2019>
  Purpose/Change: Automate VM creation
  
.EXAMPLE
    PS C:\> "Deploy.ps1"
#>

$currLocation = Get-Location

if(-not($currLocation -ilike "*Deploy VM")){
    $newLocation = Read-Host("Please enter directory of 'Deply VM' folder in order to properly run script")
    Set-Location -Path $newLocation -PassThru
}

Import-Module ..\powershell-modules\DHCP\DHCP.psm1
Import-Module ..\powershell-modules\Sharepoint\Sharepoint.psm1
Import-Module .\Modules\Numa.psm1

Function getinfo
{
	Param(
		[Parameter(Mandatory=$true)]
		[string]$vmname,
		[Parameter(Mandatory=$true)]
		$server
	)

    #Gets following info from within the VM, using previously defined session
    $VMInfo = get-vm -Name "$($vmname) - *" -Server $server | select @{N="FQDN"; E={$_.Guest.Hostname}}, @{N="Description"; E={$_.Name.ToString().Split('-')[1].Trim()}}
    $IPv4 = ((Get-VM -Name "$($vmname) - *").Guest.IPAddress | where-object {$_ -like "1*"}) -join ","
    $ipv6status = (Get-NetAdapterBinding -DisplayName "Internet Protocol Version 6 (TCP/IPv6)" | Select-Object DisplayName,Enabled | Format-List| Out-String).Trim()
    $pagefile = (Get-WMIObject Win32_PageFileSetting | Format-List | Out-String).Trim()
    $drive = (Get-HardDisk -Server $server -VM "$($item.Name) - *" | select Name, CapacityGB | Format-List | Out-String).trim() 

    #Building body of email with all info gathered above.
	     
	     "Hi, the creation stage of your request has been completed by the Company Enterprise Virtualization Team.

The following are specifications for your VM:

Host Name: $($VmInfo.FQDN)
Description: $($VMInfo.Description)
IP Address: $IPv4

The local admin access has been added to Secret Server.
The VM has been added to AD.

Please note that the VM has been created but there are still necessary steps in the workflow before the request can be completed as a whole. These steps include but are not limited to: adding the machine to SCOM, enabling backups (if client required), and completing the SLA. These steps may take some time for completion. Feel free to contact us at any time for any questions regarding your VM. 

Thank you,
	The Company Enterprise Virtualization Team `n
-------------------------------------------------------------------------------
"
	    #Get drive info
	        "$drive`n`n"
            
	    #Get pagefile 
	        "$pagefile`n`n"

	    #IPv6 status
	        "$ipv6status `n"

	    #Get list of local admin
	        #"$localadmin `n"
}

Function send-email
{
	Param(
		$vmname,
		$server
	)

    #Send email to #vm-notifications for getinfo 
    $Body = getinfo -vmname $vmname -server $server | Out-String

    $smtpserver = "smtp.company.com"
    $msg = new-object Net.Mail.MailMessage 
    $smtp = new-object Net.Mail.SmtpClient($smtpserver)

    $msg.From="name@comp.com"
    $msg.subject="VMware $vmname Build"
    $msg.To.Add("ex@amer.teams.ms")
    $msg.body = "VMware $vmname build is attached.`n`n$Body"

    $smtp.send($msg) 
}

Function Get-BackupTag{
    
    param(
        [string]$Location,
        [string]$Tag
    )

    return $value = switch($Location){
        "location1"{
            switch($Tag){
                    "No" { 
                        "None"
                        break
                    }
                    "Yes - SQL" {
                        "API"
                        break
                    }
                    "Yes - Flat Files" {
                        "API"
                        break
                    }
                    "Yes - Active Directory" {
                        "API"
                        break
                    }
                    "Yes - Oracle DB" {
                        "Oracle"
                        break
                    }
                    default {
                        "Pending"
                    }
                }
                break
            }
        "location2"{
            switch($Tag){
                "No" {
                    "NONE"
                    break
                }
                "SQL" {
                    "SQL"
                    break
                }
                "FS" {
                    "FS"
                    break
                }
                "AD" {
                    "AD"
                    break
                }
                "Oracle" {
                    "Oracle"
                    break
                }
                "API"{
                    "API"
                }
                "SP"{
                   "SP"
                   break 
                }
                default {
                    "Pending"
                }
            }
            break
        }
        default{
            "Pending"
        }      
    }
}

#region Import Data
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$ImportDialog = New-Object System.Windows.Forms.OpenFileDialog
$ImportDialog.Title = "Select List of VMs CSV"
$ImportDialog.Multiselect = $false
$ImportDialog.filter = "Comma Delimetered (*.csv)|*.csv|Text File (*.txt)| *.txt|All Files (*.*)|*.*"
$ImportDialogResult = $ImportDialog.ShowDialog()

if($ImportDialogResult -ne "OK")
{
    break
}

$CSV = Import-Csv -Path $ImportDialog.FileName | Select-Object `
        @{label="RITM"; expression={$_."request_item.number"}},
        @{label="Name"; expression={$_."variables.e45b70444faf5a00bd3c90918110c797"}},
        @{label="Description"; expression={$_."variables.88b61f674f654600bd3c90918110c7f6"}},
        @{label="VMFolder"; expression={"Create VM Staging Area"}},
        @{label="Tags"; expression={@{
            "Department" = $_."variables.21d1b62b4fca7a005af423d18110c703";
            "Backup" = Get-BackupTag -Location ($_."variables.1bf3681b4fa3620094dd97dd0210c79a") -Tag ($_."variables.cc7192024f290600db3648f18110c7a2");
            "Application" = switch($_."variables.cc7192024f290600db3648f18110c7a2") {   
                    "Yes - SQL" {
                        "SQL"
                        break
                    }
                    "Yes - Active Directory" {
                        "AD"
                        break
                    }
                    "Yes - Oracle DB" {
                        "Oracle"
                        break
                    }
                    default {
                        $null
                    }
                        
                };
            "Location" = $_."variables.1bf3681b4fa3620094dd97dd0210c79a";
            "Environment" = $_."variables.336192024f290600db3648f18110c79d"
            }
        }},
        @{label="CPU"; expression={[int]($_."variables.9eb039844f3d6200be64f0318110c786")}},
        @{label="RAM"; expression={[int]($_."variables.c1d0b9884f3d6200be64f0318110c771")}},
        @{label="Location"; expression={($_."variables.1bf3681b4fa3620094dd97dd0210c79a")}},
        @{label="Template"; expression={$_."variables.6e6192024f290600db3648f18110c76e"}},
        @{label="Cluster"; expression={$_."variables.e386a2224fec2600be64f0318110c788"}},
        @{label="ProdStatus"; expression={$_."variables.336192024f290600db3648f18110c79d"}},
        @{label="Network"; expression={if($_."variables.5193a6224fec2600be64f0318110c7bb" -ne "Other") { $_."variables.5193a6224fec2600be64f0318110c7bb" } else { $_."variables.3ad97bae4fa0a6005af423d18110c723" }}},
        @{label="Disks"; expression={
            $Disks = @()
            #Disk 1
            if($_."variables.d9df2ee24fec2600be64f0318110c7b7" -ne "")
            {
                $Disks += $_."variables.d9df2ee24fec2600be64f0318110c7b7"
            }
            #Disk 2
            if($_."variables.03b3bba64f206600be64f0318110c769" -ne "")
            {
                $Disks += $_."variables.03b3bba64f206600be64f0318110c769"
            }
            #Disk 3
            if($_."variables.47c3fba64f206600be64f0318110c7c2" -ne "")
            {
                $Disks += $_."variables.47c3fba64f206600be64f0318110c7c2"
            }
            #Disk 4
            if($_."variables.2ed3f7a64f206600be64f0318110c73e" -ne "")
            {
                $Disks += $_."variables.2ed3f7a64f206600be64f0318110c73e"
            }
            #Disk 5
            if($_."variables.c7e3fba64f206600be64f0318110c7eb" -ne "")
            {
                $Disks += $_."variables.c7e3fba64f206600be64f0318110c7eb"
            }
            $Disks}},
        @{label="Domain"; expression={$_."variables.5bc382794f431600db3648f18110c7d2"}},
        @{label="Domain_Short"; expression={$_."variables.5bc382794f431600db3648f18110c7d2".ToString().split(".")[0]}}
#endregion

#region Login
$Creds = Get-Credential -Message "Enter vCenter Credentials"
$Connections = @{}
$Locations = $CSV.Location | Select -Unique
foreach($Loc in $Locations)
{
    $server = switch($Loc)
    {
        "location1" {
            "net####.company.com"
            break
            }
        "location2" {
            "net####.company.com"
            break
        }
    }
    $Connections[$Loc] = Connect-VIServer -Server $server -Credential $Creds  -WarningAction SilentlyContinue

    while($Connections[$Loc].IsConnected -ne "True"){
        Write-Host "Error logging in to $($server)"
        $Creds = Get-Credential -Message "Enter vCenter Credentials"
        $Connections[$Loc] = Connect-VIServer -Server $server -Credential $Creds  -WarningAction SilentlyContinue
    }
}

$DHCPConnect = Connect-DHCP -Credential (Get-Credential -Message "Enter DHCP Credentials")

if($DHCPConnect -eq $False){
    "DHCP could not be connected, try again..."
    $DHCPConnect = Connect-DHCP -Credential (Get-Credential -Message "Enter DHCP Credentials")
}

#region Cloning template
$Jobs = @()
foreach($item in $CSV)
{
    #Checks to see if VM already exists
    $checkForVM = Get-VM -Name "$($item.Name) - *" -Server $server -WarningAction SilentlyContinue

    try{
   
        if($checkForVM -ne $null){
            throw "1"
        }
        
        if($item.Location -eq "location1"){
            if($item.ProdStatus -ilike "prod*" -or $item.ProdStatus -ilike "QA*")
            {
                $item | Add-Member -MemberType NoteProperty -name "ProdNonProd" -value "Production" -Force
            }
            else
            {
                $item | Add-Member -MemberType NoteProperty -name "ProdNonProd" -value "Non Production" -Force
            }
        }elseif($item.Location -eq "location2"){
            if($item.ProdStatus -ilike "prod*" -or $item.ProdStatus -ilike "QA*")
            {
                $item | Add-Member -MemberType NoteProperty -name "ProdNonProd" -value "$($item.Cluster) Production" -Force
            }
            else
            {
                $item | Add-Member -MemberType NoteProperty -name "ProdNonProd" -value "$($item.Cluster) Non Production" -Force
            }
        }

        $Parameters = @{}

        #region Teplate & OSCustomization
        $OSTranslate = @{
            "Windows Server 2016" = "2016 Template"
            "Windows Server Core 2016 Core" = "2016 Template"
            "Windows Server 2019" = "2019 Template"
            "Windows Server 2019 Core" = "2019 Template"
            "Windows Server 2012 R2" = "2012_R2 Template"
        }

        $Parameters["Template"] = Get-Template $OSTranslate[$item.Template] -Server $Connections[$item.Location]
        $Parameters["OSCustomizationSpec"] =  Get-OSCustomizationSpec -Server $Connections[$item.Location] | Where-Object { $_.name -ilike "*_$($item.domain_short)_*" -and $_.name -notlike "*VDI*" }
        #endregion

        #region Datastore
        $Datastore = switch($item.Cluster)
        {
            "location1-shared01"{
                "vSAN" 
                break
            }
            "location1-erp01"{
                "Pure"
                break
            }
            "location2C05" {
                "location2C0"
                break
            }
            "location2C" {
                "location2C01"
                break
            }
        }
        $DSTemp = Get-DatastoreCluster -Name $Datastore -Server $Connections[$item.Location] -ErrorAction SilentlyContinue
        if($null -eq $DSTemp)
        {
            $DSTemp = Get-Datastore -Name $Datastore -Server $Connections[$item.Location]
        }
        $Parameters["Datastore"] = $DSTemp
        $item | Add-Member -MemberType NoteProperty -Name "Datastore" -Value $DSTemp -Force
        #endregion

        #region ResourcePool
        $Parameters["ResourcePool"] = Get-ResourcePool -Name "$($item.ProdNonProd)" -Location (Get-Cluster -Name $item.cluster -Server $Connections[$item.Location]) -Server $Connections[$item.Location]
        #endregion

        Write-Host ('Creating {0} at {1}' -f $item.Name, $item.Location)
        $Jobs += New-VM -Name $item.Name -Location $item.VMFolder -DiskStorageFormat Thin @Parameters -Notes $item.RITM -RunAsync -Confirm:$False -Server $Connections[$item.Location]
        
    }catch{

        Write-Warning "$($item.Name) already exists in $($item.Location)"
    }
}

Write-Host "Waiting on current VM(s) to be cloned"
 
while($Jobs.State -contains "Running")
{
    Sleep 10
    $Jobs = Get-Task -Id $Jobs.Id
}
#endregion

Write-Output "VM(s) have been cloned"

#region Configure VM
foreach($item in $CSV)
{
    #region Set VM/Numa
    Write-Output "Working on '$($item.Name)'"
    $item | Add-Member -MemberType NoteProperty -Name "VM" -Value (Get-VM $item.Name)

    $item.VM = $item.VM | Set-VM -CoresPerSocket $item.CPU -NumCpu $item.CPU -MemoryGB $item.RAM -Name "$($item.Name) - $($item.Description)" -Confirm:$false

    Write-Host "Checking Numa info"
    CheckNuma -VMInfo $item -Server $Connections[$item.Location]
    #endregion

    #region Disk
    $i = 0
    foreach($disk in $item.Disks)
    {
        Write-Host "Adding drive $i into to $($item. Name)"
        New-HardDisk -CapacityGB $disk -ThinProvisioned -Datastore $item.Datastore -VM $item.VM -Server $Connections[$item.Location] | Out-Null
        $i++
    }
    #endregion

    #region Network stuffs
    $NetAdapters = $item.VM | Get-NetworkAdapter

    if($item.Location -eq "location1"){
        if(!$item.Network){
            
            $Network = Get-VirtualPortGroup -Server $Connections[$item.Location] | Out-GridView -PassThru -Title "Choose A Network For $($item.Name)"
            $NetworkInfo = $Network.Name.split("_")
            
            $VLAN = Get-DHCPVLAN -NetworkAddress $NetworkInfo[0] -NetmaskLength $NetworkInfo[1]

            Write-Host "retrieving IP Reservation from DHCP"
            $IP = Get-DHCPNextFreeIP -VLAN $VLAN
            if((Set-DHCPCreateReservation -MACAddress $NetAdapters[0].MacAddress -IPAddress $IP -Description $item.Description -VLAN $VLAN -Hostname $item.Name).StatusDescription -ne "OK")
            {
                Write-Warning -Message "Cannot create DHCP reservation"
                continue
            }

            Write-Host "Setting Network Adapter"
            $NetAdapters[0] | Set-NetworkAdapter -Portgroup $Network -Confirm:$false

        }else{
            
            $NetworkInfo = $item.Network.split("_")
            $VLAN = Get-DHCPVLAN -NetworkAddress $NetworkInfo[0] -NetmaskLength $NetworkInfo[1]

            Write-Host "retrieving IP Reservation from DHCP"
            $IP = Get-DHCPNextFreeIP -VLAN $VLAN
            if((Set-DHCPCreateReservation -MACAddress $NetAdapters[0].MacAddress -IPAddress $IP -Description $item.Description -VLAN $VLAN -Hostname $item.Name).StatusDescription -ne "OK")
            {
                Write-Warning -Message "Cannot create DHCP reservation"
                continue
            }

            Write-Host "Setting Network Adapter"
            $NetAdapters[0] | Set-NetworkAdapter -Portgroup $item.Network -Confirm:$false

        }
    }elseif($item.Location -eq "location2"){

        $VLAN = Get-DHCPVLAN -NetworkAddress "123.123.123.123" -NetmaskLength "21"

        Write-Host "retrieving IP Reservation from DHCP"
        $IP = Get-DHCPNextFreeIP -VLAN $VLAN
        if((Set-DHCPCreateReservation -MACAddress $NetAdapters[0].MacAddress -IPAddress $IP -Description $item.Description -VLAN $VLAN -Hostname $item.Name).StatusDescription -ne "OK")
        {
            Write-Warning -Message "Cannot create DHCP reservation"
            continue
        }

        Write-Host "Setting Network Adapter"
        $NetAdapters[0] | Set-NetworkAdapter -Portgroup "EX# VLAN" -Confirm:$false
    }
    #endregion

    if($null -ne $item.VM){
        Start-VM -VM $item.VM
    }

    do{
        sleep -Seconds 5
        $VM = Get-VM $item.VM | Select-Object Guest
        $powerstate = $VM.Guest.State

    }while($powerstate -ne "running")

    #Check if domain joined before calling Set VM script
    Write-Host "Waiting on '$($item.Name)' to join the Domain"
    while((Get-VM $item.VM).Guest.HostName -notlike "*.$($item.Domain)")
    {
        Sleep 10
    }
    Write-Host "'$($item.Name)' has joined the Domain" 

    #region Folder move
    
    <#$Folder = $item.Tags.Department + "_Test"

    $FolderList = Get-Folder -Server $Connections[$item.Location] -Type VM

    #Check to see if folder exists in vCenter. If not, create new folder and move VM into folder.
    if($FolderList.Name -inotcontains $Folder){
        Write-Warning "$Folder does not exist in vCenter, creating new Folder..."
        
        $FolderCheck = New-Folder -Name $Folder -Server $Connections[$item.Location].Name -Location ((Get-VM -Name "$($item.Name) - *" -Server $Connections[$item.Location].Name) | Get-Datacenter -Name "location1-w01dc" | Get-Folder -Name vm)
        
        if($FolderCheck){
            Move-VM -VM "$($item.Name) - *" -InventoryLocation (((Get-VM -Name "$($item.Name) - *" -Server $Connections[$item.Location].Name) | Get-Datacenter) | Get-Folder -Name $Folder)
        }elseif(!$FolderCheck){
            Write-Warning "Folder could not be created, manually create '$Folder' in $($Connections[$item.Location].Name.ToString().Split('.')[0]) within $($item.Cluster) under 'VMs and Templates'."
            continue
        }
    }else{
        Write-Host "Moving $($item.Name) into $Folder Folder..."
        Move-VM -VM "$($item.Name) - *" -InventoryLocation (((Get-VM -Name "$($item.Name) - *" -Server $Connections[$item.Location].Name) | Get-Datacenter) | Get-Folder -Name $Folder) -WhatIf
    }#>
    #endregion
    
    #region AD section
    #Needs fixing
    <#$objGroup = Get-ADGroup -LDAPFilter "(name=ex*)" -SearchScope Subtree -SearchBase "ou=Group,dc=net,dc=comp,dc=com" | select name
    $GroupName = "Group_" + $item.Tags.Department + "_" + $item.homeDepart
    
    #Check AD to see if Securtiy Group exists. If not, create new group under OU and move VM into Group
    if(-not($objGroup.name -ilike $GroupName)){
        Write-Host "Need to create new Security Group with name: " -NoNewline
        Write-Host $GroupName -ForegroundColor Cyan
        New-ADGroup -Name $GroupName -SamAccountName $GroupName -GroupCategory Security -GroupScope Global -DisplayName $GroupName -Path "ou=Group,dc=net,dc=comp,dc=com" -Description "Department: $($item.Tags.Department), Home Department: $($item.homeDepart)" -WhatIf
    }#>
    #endregion

    #region Tags
    foreach($Category in $item.Tags.Keys)
    {
        if($item.Tags.$Category)
        {
            if($item.Location -ne "location2" -or $Category -ne "Application"){
                $Tag = Get-Tag -Category $Category -Name $item.Tags.$Category -Server $Connections[$item.Location]
                if($null -eq $Tag)
                {
                    $Tag = New-Tag -Category $Category -Name $item.Tags.$Category -Server $Connections[$item.Location]
                }
                New-TagAssignment -Entity $item.VM -Tag $Tag -Server $Connections[$item.Location]
                Write-Host "$($Tag.Name) has been added"
            }
        }
    }
    #endregion    

    #region reboot
    write-Host "Rebooting $($item.VM)..."
    Restart-VM -VM (Get-VM -Name "$($item.Name) - *" -Server $server) -Confirm:$false | Out-Null

    do{
        sleep -Seconds 5
        $VM = Get-VM "$($item.Name) - *" | Select-Object Guest
        $powerstate = $VM.Guest.State
    }while($powerstate -ne "running")
    #endregion

    #Function to send email to Teams notif channel (Get working creds)
    send-email -vmname $item.Name -server $Connections[$item.Location].Name -cred $Creds
    #endregion

    #region Config
    Write-Host "Entering Config stage"
    & '..\VMware_vRealize_Automation\Powershell\Windows Post-OS Config\Set-VM.ps1' -vmname $item.Name -location VMware
    #endregion
}
#endregion