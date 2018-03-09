$global:pcip
function CheckHost{ 
    param ([ValidateNotNullOrEmpty()]
    $compname  )
    $ping = gwmi Win32_PingStatus -ErrorAction Stop -filter "Address='$compname'" 
    $ping.ProtocolAddress.GetType()
    if($ping.StatusCode -eq 0){$pcip=$ping.ProtocolAddress; return 1} 
    else{return 0} 
} 
function cut-string   
{  
    param(  
        [String]$str,  
        [String]$start,  
        [String]$end  
    )  
    return $str.substring($str.indexof($start),$str.indexof($end)-$str.indexof($start))  
}
function GetLogin{
    param ([ValidateNotNullOrEmpty()]
    $compname  )
    $loginInfo=Get-WinEvent -computerName $compname -logname security -maxevents 100| ? {$_.id -eq 4624} | select timecreated,message 
    $str=$loginInfo.message
    $conbin=''
    $time=$loginInfo.timecreated
    if($str.Count -eq 1){
        try{
             $WWW=cut-string $str 'Account Name:' 'Account Domain:'
        }catch{
            "String cut error"
            continue
         }
        $WWW=$WWW -replace 'Account Name:', ""
        $WWW=$WWW -replace "		" ,""
        $WWW=$WWW -replace "	" ,""
        $WWW=$WWW -replace "`r" ,""
        $WWW=$WWW -replace "`n" ,""
        $conbin=$conbin+$WWW+","+$time[$i]+"`r`n"
    }
    else{
       try{
        $WWW=cut-string $str[0] 'Account Name:' 'Account Domain:'
      }catch{
        "String cut error"
        continue
      }
    $WWW=$WWW -replace 'Account Name:', ""
    $WWW=$WWW -replace "		" ,""
    $WWW=$WWW -replace "	" ,""
    $WWW=$WWW -replace "`r" ,""
    $WWW=$WWW -replace "`n" ,""
    $conbin=$conbin+$WWW+","+$time[0]+"`r`n"
    $conbin=$conbin -replace "	" ,""
    }
    $conbin=$conbin -replace "	" ,""
    return $conbin
}
function WantNext{
    $MenuSelection = Read-Host "Do you want to try again?[y/n]" 
    return $MenuSelection
}
$s=Get-Content "$PSScriptRoot\scs_hostnames.txt"
Clear-Host
while(1 -eq 1)
{
    $s|ForEach-Object{
        $index=[array]::indexof($s,$_)
        "$_..............[$index]"
    }
    "All..............[-1]"
    "Exit..............[-2]"
    $MenuSelection = Read-Host "Enter Selection" 
    Clear-Host
    if($MenuSelection -eq -1){
        $s|ForEach-Object {
            [String] $result=$compname+"`r`n"
            $compname=$_
            $state=CheckHost $_
            if($state -eq 1){
                "Connect to $compname"
                try{
                    $result+="SerialNumber: "+[String](gwmi -computer $compname -ErrorAction Stop Win32_BIOS | Select-Object SerialNumber).SerialNumber+"`r`n"
                    $rs=gwmi -ErrorAction Stop -computer $compname Win32_Printer | Select-Object DeviceID,DriverName, PortName
                    $result+="Printer Info:`r`n"
                    $rs|ForEach-Object {
                        $result+="  DeviceID: "+[String]$_.DeviceID+"`r`n       DriverName: "+[String]$_.DriverName+"`r`n       PortName: "+[String]$_.PortName+"`r`n"
                    }
                    $result+="Current User Name: "+(gwmi -ErrorAction Stop -computer $compname Win32_ComputerSystem).Username+"`r`n"
                    $rs=(gwmi -ErrorAction Stop -computer $compname Win32_OperatingSystem)
                    $result+="OS name:"+$rs.Caption+" "+$rs.OSArchitecture+"`r`n"
                    $result+=(GetLogin $compname)+"`r`n"
                    "Get Info"
                    $result
                    $date=Get-Date -UFormat "%Y.%m.%d"
                    $result|Out-File "$PSScriptRoot\$date"+"_"+"$compname.txt"
                }catch{
                    "Catch Error"
                    $result+="$compname connection Error"+"`r`n"
                    $_|Out-File "$PSScriptRoot\log.txt"
                    continue
                }
                
            }
            else{
                "Can not find $compname"
            }
            
        }
    }
    elseif($MenuSelection -eq -2){
        exit
    }
    else{
        $compname=$s[$MenuSelection]
            $state=CheckHost $_
            if($state -eq 1){
                "Connect to $compname"
                try{
                    $result+=$compname+": "+"`r`n"
                    $result+="SerialNumber: "+[String](gwmi -computer $compname -ErrorAction Stop Win32_BIOS | Select-Object SerialNumber).SerialNumber+"`r`n"
                    $rs=gwmi -computer $compname -ErrorAction Stop Win32_Printer | Select-Object DeviceID,DriverName, PortName
                    $result+="Printer Info:`r`n"
                    $rs|ForEach-Object {
                        $result+="  DeviceID: "+[String]$_.DeviceID+"`r`n       DriverName: "+[String]$_.DriverName+"`r`n       PortName: "+[String]$_.PortName+"`r`n"
                    }
                    $result+="Current User Name: "+(gwmi -computer $compname -ErrorAction Stop Win32_ComputerSystem).Username+"`r`n"
                    $rs=(gwmi -computer $compname -ErrorAction Stop Win32_OperatingSystem)
                    $result+="OS name:"+$rs.Caption+" "+$rs.OSArchitecture+"`r`n"
                    $result+=(GetLogin $compname)+"`r`n"
                    "Get Info"
                    $result
                    $date=Get-Date -UFormat "%Y.%m.%d"
                    $result|Out-File "$PSScriptRoot\$date"+"_"+"$compname.txt"
                }catch{
                    "Catch Error"
                    $result+="$compname connection Error"+"`r`n"
                    $_|Out-File "$PSScriptRoot\log.txt"
                }
            }
            else{
                "Can not find $compname"
            }
            Read-Host "Press ENTER to continue" 
    }
    Clear-Host
}