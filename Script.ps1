#region - SETTING EXECUTION POLICY 

    #Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    #Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope [Machine, User, Process, CurrentUser, LocalMachine]
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

#endregion

#region - INSTALLING EXCEL MODULE IF NOT AVAILABLE

    If(-not(Get-InstalledModule ImportExcel -ErrorAction silentlycontinue)){

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
        Install-Module -Name ImportExcel -RequiredVersion 5.4.0  -Confirm:$False -Force
        #Install-Module -Name ImportExcel -RequiredVersion 7.1.1  -Confirm:$False -Force

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
    $DataFile = "Data.xlsx"
    $Data = $DataPath + $DataFile

    $DMZServerNamesTXT = "dmzserverhostnames.txt"
    $DMZServerNames = $DataPath + $DMZServerNamesTXT

    $LWGServerNamesTXT = "workgroupserverhostnames.txt"
    $LWGServerNames = $DataPath + $LWGServerNamesTXT

    $domainName = ""


#endregion

#region - PASSWORD FOR LOCALGROUP SERVERS

# 'notactuallymypassword' | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "Password.txt"
# $PWord = Get-Content "C:\Password.txt" | ConvertTo-SecureString

$PWord = ConvertTo-SecureString -String "P@sSwOrd" -AsPlainText -Force

#endregion

#region - SERVICE ACCOUNT FOR DMZ SERVERS

    $servPass = ConvertTo-SecureString -String "P@sSwOrd" -AsPlainText -Force
    $servUser = ""

#endregion

Write-Output "###############################################################################" | TimeStamp | Out-File $logData -Append
Write-Output "                                STARTED Script                                 " | TimeStamp | Out-File $logData -Append
Write-Output "###############################################################################" | TimeStamp | Out-File $logData -Append


#region - GETTING WMI DATA FOR DMZ SERVERS

    Write-Output "Write-Output Started WMI DATA GATHERING FOR DMZ SERVERS" | TimeStamp | Out-File $logData -Append

    # retrieving dmz servers from a file
    $dmzservers =  Get-Content $DMZServerNames

    $ServCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $servUser, $servPass 

    foreach ($dmzserverhostname in $dmzservers) {

        Write-Output "Getting WMI data. Processed : $dmzserverhostname" | TimeStamp | Out-File $logData -Append

        try {

            Get-WmiObject `
            -ComputerName $dmzserverhostname `
            -Authority "ntlmdomain:$domainName" `
            -Credential $ServCredential `
            -Query "select * from AntiMalwareHealthStatus" `
            -Namespace "root\Microsoft\SecurityClient" `
            -ErrorAction Stop `
            | Select-object PScomputerName, version, AntivirusSignatureVersion, AntiVirusSignatureUpdateDateTime,AntivirusEnabled `
            | Export-Excel  -Append  -workSheetName "WMI-Targets" -Path $Data

        }

        catch {

            Write-Output "ERROR.....There maybe an issue with this DMZ Server : $dmzserverhostname" | TimeStamp | Out-File $logData -Append 
        
        } 
            
    } # end foreach

    Write-Output "Finished WMI DATA GATHERING FOR DMZ SERVER" | TimeStamp | Out-File $logData -Append

    Write-Output "" | Out-File $logData -Append
    Write-Output "" | Out-File $logData -Append

#endregion


#region - GETTIING WMI DATA FOR LOCAL WORKGROUP SERVERS

    Write-Output "Write-Output Started WMI DATA GATHERING FOR LOCAL WORKGROUP SERVERS" | TimeStamp | Out-File $logData -Append

    # retrieving localgroup server from a file
    $workgroupservers =  Get-Content $LWGServerNames

    foreach ($workgroupservername in $workgroupservers) {

        Write-Output "Getting WMI data. Processed : $workgroupservername" | TimeStamp | Out-File $logData -Append

        $User = "$workgroupservername\Administrator"
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord 
        #Credential  = New-Object -TypeName System.Management.Automation.PSCredential($User, $PWord) 

        try {

            Get-WmiObject `
            -ComputerName $workgroupservername `
            -Credential $Credential `
            -Query "select * from AntiMalwareHealthStatus" `
            -Namespace "root\Microsoft\SecurityClient" `
            -ErrorAction Stop `
            | Select-object PScomputerName, version, AntivirusSignatureVersion, AntiVirusSignatureUpdateDateTime,AntivirusEnabled `
            | Export-Excel  -Append  -workSheetName "WMI-Targets" -Path $Data

        }

        catch {

            Write-Output "ERROR.....There maybe an issue with this Local WorkGroup Server : $workgroupservername" | TimeStamp | Out-File $logData -Append 
        
        } 

    } 

    Write-Output "Finished WMI DATA GATHERING FOR LOCALGROUP SERVER" | TimeStamp | Out-File $logData -Append

    Write-Output "" | Out-File $logData -Append
    Write-Output "" | Out-File $logData -Append

#endregion

Write-Output "###############################################################################" | TimeStamp | Out-File $logData -Append
Write-Output "                                FINISHED Script                                " | TimeStamp | Out-File $logData -Append
Write-Output "###############################################################################" | TimeStamp | Out-File $logData -Append

Write-Output "" | Out-File $logData -Append
Write-Output "" | Out-File $logData -Append