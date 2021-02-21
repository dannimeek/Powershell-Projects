##################################################################################################   
        # NB: Run the Execution Policy command directly in the powershell console before 
        #     running the whole script.
##################################################################################################

#region - SETTING EXECUTION POLICY 
 
    # RUN THIS : Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

    <# 
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope [Machine, User, Process, CurrentUser, LocalMachine]
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    #>

#endregion

#region - INSTALLING EXCEL MODULE IF NOT AVAILABLE

    If(-not(Get-InstalledModule ImportExcel -ErrorAction silentlycontinue)){

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
        Install-Module -Name ImportExcel -RequiredVersion 5.4.0  -Confirm:$False -Force

    }

#endregion

#region - TIMESTAMPS FOR LOGGING

    # TIMESTAMP FILTER to Add a timestamp INSIDE a LOG FILE
    filter timestamp {"$(Get-Date -Format G): $_"} # TIMESTAMP for Log FILENAME

#endregion 

#region - SOME VARIABLE INITALIZATION

    $logDataPath = ""
    $logDataFile = "LogCollection.txt"
    $logData = $logDataPath + $logDataFile

    $DataPath = ""
    $DataFile = "LoggedOnInformation.xlsx"
    $Data = $DataPath + $DataFile

    # Text file to contain live computers from Ping result
    $Pingable_IPAddresses = "" 

    # Range to Ping ( 1 to 254)
    $lastOctet = 1..254

    # To hold the first 3 octet of the IP Subnet (eg: 192.168.0)
    $ipadd = "" 

    $dataObject = New-Object -TypeName psobject

#endregion

Write-Output "###############################################################################" | TimeStamp | Out-File $logData -Append
Write-Output "                                STARTED Script                                 " | TimeStamp | Out-File $logData -Append
Write-Output "###############################################################################" | TimeStamp | Out-File $logData -Append

#region - SWEEPING LIVE (PINGABLE) IP ADDRESSES 

    Write-Output "Started SWEEPING IP ADDRESSES" | TimeStamp | Out-File $logData -Append

    foreach ($octet in $lastOctet){

        if ((Test-Connection $ipadd.$octet -Count 1 -Delay 1 -ErrorAction Ignore)){

            Write-Output "$ipadd.$octet" | Out-File  $Pingable_IPAddresses -Append
        
        }

    }

    Write-Output "Finished SWEEPING IP ADDRESSES" | TimeStamp | Out-File $logData -Append

#endregion

#region - WMI QUERY FOR USERNAME, COMPUTERNAME, OPERATION SYSTEM AND IPADDRESS, AND EXPORTING TO EXCEL FILE

    Write-Output "Started WMI QUERY" | TimeStamp | Out-File $logData -Append

    $Computers = Get-Content $Pingable_IPAddresses
    
    foreach($ip in $Computers){

        $user = Invoke-Command $ip {(Get-WmiObject -Class win32_Process -Filter 'Name="explorer.exe"').GetOwner().User} -ErrorAction SilentlyContinue
        
        $compname = Invoke-Command $ip {(Get-WmiObject -Class win32_Process -Filter 'Name="explorer.exe"').CSName} -ErrorAction SilentlyContinue
        
        $fullName = Get-ADUser $user | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
        
        $OSinfo = Get-ADComputer -Identity $compname -Properties * | Select-Object -ExpandProperty Operatingsystem
 
        $IPinfo = Get-ADComputer -Identity $compname -Properties * | Select-Object -ExpandProperty IPv4Address

        # Write-Output "$fullName is currently logged on to $compname (Running $OSinfo)"

        $dataObject | Add-Member -MemberType NoteProperty -Name LoggedOnUser -Value $fullName
        $dataObject | Add-Member -MemberType NoteProperty -Name ComputerLoggedOnTo -Value $compname
        $dataObject | Add-Member -MemberType NoteProperty -Name RunningOperationSystem -Value $OSinfo
        $dataObject | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPinfo
        
        $dataObject | Export-Excel  -Append  -workSheetName "LoggedOnInformation" -Path $Data
    

    }

    Write-Output "Finished WMI QUERY" | TimeStamp | Out-File $logData -Append

    Write-Output "LoggedOnInformation saved to $DataFile Successfully" | TimeStamp | Out-File $logData -Append

#endregion


Write-Output "###############################################################################" | TimeStamp | Out-File $logData -Append
Write-Output "                                FINISHED Script                                " | TimeStamp | Out-File $logData -Append
Write-Output "###############################################################################" | TimeStamp | Out-File $logData -Append
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        