<#
Powershell adaptation of narra macro

#>

$macro_dir                 = "S:\macro.me\New Release"
$programs_dir              = "\\ireland\programs"
$programs_drive            = "S:"
$grp_name                  = "Narra"

$title                     = “Multi-Edit for $grp_name”
$host.ui.RawUI.WindowTitle = $title


$str_errMissing   = "We can't find a Multi-Edit directory.`nPlease install Multi-Edit 2006 to proceed with the installation.`n`n"
$str_errAdmin     = "As per checking, you don't have any admin priveledges. Please check if you are a valid user for this PC."            
$str_errProgDrive = "You don't have any mapping to $programs_drive,`nso we have made one for you.`n`n"
$str_hdGroup      = "Please choose a group"
$str_hdGeneral    = "Updating global configuration files"
$str_hdSetup      = "Setting up Multi-Edit for"
$str_msConfigured = "$grp_name has been configured in this PC"

$bool_installed   = "F"

<#
******************************
  1. Gathering Information
******************************
#>

if([System.IntPtr]::Size -eq 8){$os_arch = "64"} else{$os_arch = "32"}
if($os_arch -eq "64") {$multi_dir = "$env:SystemDrive\Program Files (x86)\Multi-Edit 2006"}
else {$multi_dir = "$env:SystemDrive\Program Files\Multi-Edit 2006"}

if(([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")) -eq $false){
    Write-Output $str_errAdmin
    CMD /c PAUSE
    exit
}

if( -not (Test-Path -Path $multi_dir -PathType Container)){
   Write-Output $str_errMissing
   CMD /c PAUSE
   exit
}

if (-not (Test-Path -Path $programs_drive -PathType Container)) {
   (new-object -com WScript.Network).MapNetworkDrive($programs_drive,$programs_dir)
   Write-Output $str_errProgDrive
   CMD /c PAUSE
   cls
}

if(Test-Path -Path "$multi_dir\Database" -PathType Container){
    $bool_installed = "T"
}


cls

<#
******************************
  2. Updating Configurations
******************************
#>


function gen-config {
Write-Host "
 1. $str_hdGeneral"

    Get-ChildItem -Path "$macro_dir\Config\*" -Include *.* | Copy-Item -Destination "$multi_dir\Config"
    Get-ChildItem -Path "$macro_dir\Macro\common\*" -Include *.mac | Copy-Item -Destination "$multi_dir\Mac"
    if( -not (Test-Path -Path "$multi_dir\Database" -PathType Container)){
        [void](New-Item -ItemType directory -Path "$multi_dir\Database")
    }
    Copy-Item "$macro_dir\Database\westlaw.db" "$multi_dir\Database\westlaw.db"
}


function update-config {

    Set-ItemProperty "$multi_dir\Config\*.db" -name IsReadOnly -value $false
    Set-ItemProperty "$multi_dir\Config\*.tpt" -name IsReadOnly -value $false
    
    #Get-ChildItem -Path "$multi_dir\Config\*.db" -Recurse -File | % { $_.IsReadOnly=$False }
    #Get-ChildItem -Path "$multi_dir\Config\*.tpt" -Recurse -File | % { $_.IsReadOnly=$False }
    
    $files = @("auxdic.txt","auxdic2.txt")
    For ($i=0; $i -lt $files.Length; $i++) {
        $temp = -join("$multi_dir\Utils\",$files[$i])
        if(Test-Path -Path $temp -PathType Leaf){ Get-ChildItem $temp | Remove-Item -Force }
        
        $temp = -join("$multi_dir\Config\",$files[$i])
        if(Test-Path -Path $temp -PathType Leaf){ Get-ChildItem $temp | Remove-Item -Force }
    }
}

function check-config {
    if( -not (Test-Path -Path "$env:public\Desktop\Multi-Edit 2006 NARRA.lnk" -PathType Leaf)){
        Get-ChildItem -Path "$macro_dir\icons\cos" -Recurse | Copy-Item -Destination "$env:public\Desktop"
    }
}

function remove {

    $files = @("LSS.TPT","Meconfig.db","Sgml.tpt","Wcmdmap.db","WESTLAW.TPT")
    For ($i=0; $i -lt $files.Length; $i++) {       
        $temp = -join("$multi_dir\Config\",$files[$i])
        if(Test-Path -Path $temp -PathType Leaf){ Get-ChildItem $temp | Remove-Item -Force }
    }
    
    $files = @("NARRA.DAT","westlaw.db")
    For ($i=0; $i -lt $files.Length; $i++) {       
        $temp = -join("$multi_dir\Database\",$files[$i])
        if(Test-Path -Path $temp -PathType Leaf){ Get-ChildItem $temp | Remove-Item -Force }
    }
    
    $files = @("common.mac","csnysck.mac","cspadcck.mac","csprck.mac","galesys.mac","lsssys.mac","MainLib.mac","narrasys.mac","sgmlsys.mac",
                "Spell.mac","wciteex.mac","westlaw.mac","westlaw_20130605.mac","westlaw_20130910.mac","westlaw_20140325.mac","westlaw_20140328.mac",
                "westlaw_20140502.mac","westlaw_20140603.mac","westlaw_20140730.mac","westlaw_20140904.mac","westlaw_20160302.mac","westlaw_20160502.mac",
                "westlaw_20160802.mac","westlaw1.mac")
    For ($i=0; $i -lt $files.Length; $i++) {       
        $temp = -join("$multi_dir\Mac\",$files[$i])
        if(Test-Path -Path $temp -PathType Leaf){ Get-ChildItem $temp | Remove-Item -Force }
    }
    
    $files = @("Sencor.dic")
    For ($i=0; $i -lt $files.Length; $i++) {       
        $temp = -join("$multi_dir\Utils\",$files[$i])
        if(Test-Path -Path $temp -PathType Leaf){ Get-ChildItem $temp | Remove-Item -Force }
    }
    
    $files = @("LSS Court Orders","LSS","NARRA")
    For ($i=0; $i -lt $files.Length; $i++) {      
        $temp = -join("$env:public\Desktop\Multi-Edit 2006 ",$files[$i],".lnk")
        #Write-Host $temp
        if(Test-Path -Path $temp -PathType Leaf){ Get-ChildItem $temp | Remove-Item -Force }
    }
    
    if(Test-Path -Path "$multi_dir\narra.ini" -PathType Leaf){ Get-ChildItem "$multi_dir\narra.ini" | Remove-Item -Force }
    Remove-Item "$multi_dir\Database"
}


function setup{
    Write-Host "
 ----------------------------------------------

 $str_hdSetup Narra
       
 ----------------------------------------------"

    gen-config
    
    Write-Host " 2. Installing $grp_name configuration files"
    
    #narra run-first    
    Get-ChildItem -Path "$macro_dir\Macro\narra\*" -Include *.mac | Copy-Item -Destination "$multi_dir\Mac"
    Copy-Item "$macro_dir\Utils\narra.dic" "$multi_dir\Utils\Sencor.dic"
    Copy-Item "$macro_dir\narra.ini" "$multi_dir\narra.ini"
    Get-ChildItem -Path "$macro_dir\Database\*" -Include narra.* | Copy-Item -Destination "$multi_dir\Database"
    
    #narra update
    Write-Host " 3. Running updates`n`n"
    update-config

    Write-Host " Installation Complete. `n"               
    CMD /c PAUSE  
}


do{

if (Test-Path -Path $multi_dir\narra.ini -PathType Leaf) {

   cls

   Write-Host "

 NN   NN    AAA    RRRRRR   RRRRRR     AAA   
 NNN  NN   AAAAA   RR   RR  RR   RR   AAAAA  
 NN N NN  AA   AA  RRRRRR   RRRRRR   AA   AA 
 NN  NNN  AAAAAAA  RR  RR   RR  RR   AAAAAAA 
 NN   NN  AA   AA  RR   RR  RR   RR  AA   AA
 
 ----------------------------------------------

 1. Reinstall $grp_name-SETUP
 2. Add desktop icons
 3. Install fonts
 4. Reinstall macros
   
 0. Uninstall SENCOR configuration

 X. Close

 ----------------------------------------------`n"
   $option = Read-Host -Prompt ' Select option'


   if($option -eq "0"){
       cls
       Write-Host "
 ----------------------------------------------
       
 0. Removing SENCOR configuration
       
 ----------------------------------------------"
    
       remove
    
       Write-Host "`n`n Deletion complete.`n`n "
    
       CMD /c PAUSE 
       break
   }

   elseif($option -eq "1"){
        cls
        <#
        Write-Host "
 ----------------------------------------------
       
 1. Reinstalling $grp_name-SETUP
       
 ----------------------------------------------"
 #>
        setup

        Write-Host "`n`n Installation complete.`n`n "
        CMD /c PAUSE 

        break
   }

   elseif($option -eq "2"){
        cls
        Write-Host "
 ----------------------------------------------
       
 2. Adding Multi-Edit Desktop icons for LCP
       
 ----------------------------------------------"
        check-config

        Write-Host "`n`n Installation complete.`n`n "
        CMD /c PAUSE 

        break
   }

   elseif($option -eq "3"){
        cls
        Write-Host "
 ----------------------------------------------
       
 3. Installing Multi-Edit fonts
       
 ----------------------------------------------"
        check-config

        Write-Host "`n`n Installation complete.`n`n "
        CMD /c PAUSE 

        break
   }

   elseif($option -eq "4"){
        cls
        Write-Host "
 ----------------------------------------------
       
 4. Installing $grp_name Macros
       
 ----------------------------------------------"
        Get-ChildItem -Path "$macro_dir\Macro\common\*" -Include *.mac | Copy-Item -Destination "$multi_dir\Mac"
        Get-ChildItem -Path "$macro_dir\Macro\narra\*" -Include *.mac | Copy-Item -Destination "$multi_dir\Mac"

        Write-Host "`n`n Installation complete.`n`n "
        CMD /c PAUSE 

        break
   }   
}

else{
    cls
    setup
    check-config
    CMD /c PAUSE
    break
}


}until ($option -eq 'X')

exit