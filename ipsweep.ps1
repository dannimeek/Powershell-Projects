
 $lastOctet = 1..254

 $ipadd = "2.2.2"

    foreach ($octet in $lastOctet){
 
    # Test-NetConnection "$ipadd.$octet" -InformationLevel Quiet

    # [bool]$ping = [bool]$(Test-Connection "$ipadd.$octet" -Count 1 -Delay 1 -ErrorAction Ignore )

    [bool]$ping = Test-NetConnection "$ipadd.$octet" -InformationLevel Quiet -ErrorAction SilentlyContinue -WarningAction Ignore

    if($ping){

        Write-Output "$ipadd.$octet" >> "~\Desktop\ip.txt"

        }

    }



#[bool]$ping = [bool]$(Test-Connection 10.10.62.62 -Count 1 -Delay 1 -ErrorAction SilentlyContinue)
#[bool]$ping = [bool]$(Test-Connection 10.10.62.1 -Count 1 -Delay 1 -ErrorAction Ignore)
