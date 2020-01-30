<#
    .Synopsis
    Generate a report for a App-V Management Server and the Packages 
    .DESCRIPTION
    Generate a report for a App-V Management Server and the Packages. 
    The Verbose switch describes in detail what happens in the script
    .Parameter StartBrowser
    Run the Browser after the Report is generated
    .Parameter StartBrowser
    start a webbrowser to view the report
    .Parameter FullReport
    Create a detail report
    .Parameter ExtarctIcon
    Extract images from the appv file (it takes a little longer)
    .Parameter OutputPath 
    Instead of the temporary folder you can also specify a custom output file

    .EXAMPLE
    Test-AppVServer.ps1 -StartBrowser -FullReport -ExtractIcon -OutputPath c:\temp\myReport.html
    .INPUTS
    
    .OUTPUTS
    
    .NOTES
    Andreas Nick, 2019/2020
    Special Thanks to Thorsten Enderline @endi24 for testing and support
    Free for use at your own risk (MIT License)
    
    .COMPONENT
    Only run on AppVServer
    AppVToolsAndAnalysis

    .LINK
    https://www.software-virtualisierung.de

#>

[CmdletBinding()]
param(
  [Switch] $StartBrowser,
  [Switch] $FullReport,
  [Switch] $ExtractIcon,
  [String] $OutputPath = "$env:temp\$(get-Date -Format 'yyyymmdd-hhmmss')_AppVServerReport.html"

)

#Test for Management Server      
if(-not (get-module -ListAvailable AppVServer)){
  Write-Host "The script can only work on an App_V Management Server. The AppVServer module is required for execution." -ForegroundColor Yellow
  Throw "Missing AppV Management Server"
}


#Test for AppVForcelets
if(-not (get-module -ListAvailable AppVForcelets)){
  Write-Verbose "PowerShell module AppVForcelets not in the module folder"
  Write-Verbose "You can install it with Install-Module AppVForcelets"
  Write-Verbose "We test for a local module folder"
  
  if(Test-Path -Path "$PSScriptRoot\AppVForcelets"){
     Import-Module ('{0}\AppVForcelets' -f $PSScriptRoot) 
  } else {
     Write-Host "PowerShell module AppVForcelets not found" -ForegroundColor Yellow
     Write-Host "You can install it with Install-Module AppVForcelets -scope CurrentUser" -ForegroundColor Yellow
     Throw "Module AppVForcelets not found use -verborde for more informations"
  }
} else 
{
  Import-Module AppVForcelets
}

#

Import-Module AppVServer


Function Test-HTTPPath{
  [CmdletBinding()]
  param($path)
  #$ErrorActionPreference = "SilentlyContinue"
  $r = [System.Net.WebRequest]::Create("$Path")
  $r.UseDefaultCredentials = $true 
  $SC = $r.GetResponse().StatusCode
  if($SC -eq "OK"){ $HTTPStatus = "True" } 
  else { $HTTPStatus = "False" }
  return $HTTPStatus
}

function Test-AppVServerPackageges{
  [CmdletBinding()]
  param()

  $Packages = Get-AppvServerPackage | Sort-Object -Property Name  
  $count=1
  $Results_TAVPP = New-Object System.Collections.ArrayList

  foreach($Package in $Packages ){
    
    Write-Verbose $("Analyse App-V Package " + $Package.Name) 

    $Result_TAVPP = New-Object System.Object
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name ID -Value  $Package.Id
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name Count -Value  $count
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name PackageName -Value $Package.Name
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name PackageVersion -Value $Package.Version
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name PackageID -Value $Package.PackageGuid
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name VersionID -Value $Package.VersionGuid
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name ConnectionGroups -Value ($Package | Select-Object -Property ConnectionGroups).ConnectionGroups
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name Enabled -Value $Package.Enabled
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name Entitlements -Value $Package.Entitlements

    $Path = $Package.PackageUrl

    if(($Path.StartsWith("http")) -eq $true){
      $TPP = Test-HTTPPath "$Path"
    }

    else {
      $TPP = Test-Path "$Path"
    }

    if($TPP -eq $false){
      $Result_TAVPP | Add-Member -MemberType NoteProperty -Name PackagePath -Value $Path
      $Result_TAVPP | Add-Member -MemberType NoteProperty -Name PackageTestPath -Value "False"
      Write-Verbose $("Package Path not found " + $Package.Name)
      $Result_TAVPP | Add-Member -MemberType NoteProperty -Name ManifestInfo -Value $null
    }
    elseif($TPP -eq $true){
      $Result_TAVPP | Add-Member -MemberType NoteProperty -Name PackagePath -Value $Path
      $Result_TAVPP | Add-Member -MemberType NoteProperty -Name PackageTestPath -Value "True"
          
      Write-Verbose $("Extract AppXManifest from Package " + $Package.Name)
      $Result_TAVPP | Add-Member -MemberType NoteProperty -Name ManifestInfo -Value (Get-AppVManifestInfo $Path)
      #Extract Icons
      if($ExtractIcon -and $Result_TAVPP.ManifestInfo.HasShortcuts ){
        Write-Host "Extract Icons from $Path" -ForegroundColor Yellow
        $IconList =$Result_TAVPP.ManifestInfo.Shortcuts
            
        #Write-Host $($IconList) -ForegroundColor Cyan
            
        $result = Get-AppVIconsFromPackage -Path $Path -Iconlist @($IconList) -ImageType png
        $Result_TAVPP | Add-Member -MemberType NoteProperty -Name Icons -Value  $result
      } else{
        $Result_TAVPP | Add-Member -MemberType NoteProperty -Name Icons -Value  $null
      }
    }
        
    #Get Server Deployment COnfig
    [xml]$DepConfig = $package |  Get-AppvServerPackageDeploymentConfiguration 
    $DepConfig.save("$env:TEMP\TempAppVServerReport.xml" )
    $Result_TAVPP | Add-Member -MemberType NoteProperty -Name DeploymentConfigInfo -Value (Get-AppVDeploymentConfigInfo "$env:TEMP\TempAppVServerReport.xml")
        
    #Get DeploymentConfig form the App-V folder, if there is one
    $DepConfigPath = $Path -replace '\.appv','_DeploymentConfig.xml'
    if(Test-Path $DepConfigPath){
      $Result_TAVPP | Add-Member -MemberType NoteProperty -Name Disk_DeploymentConfigInfo -Value (Get-AppVDeploymentConfigInfo $DepConfigPath)
    } else
    {
      $Result_TAVPP | Add-Member -MemberType NoteProperty -Name Disk_DeploymentConfigInfo -Value $null
    }
        
        
    $Results_TAVPP += $Result_TAVPP
        
    $count++
  }
  return $Results_TAVPP
}


#Little Helper for a custom vertical html Table
function Add-ReportHtml{
  param(
    [string] $Titel="Title",
    [string] $Commend ="Commend",
    [psobject] $InformationObject
  )
    
  $Report = @"
<br><br>
<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">$Titel</span>
<br><br>
$Commend
<br><br>
<table width="40%" style="border-collapse: collapse; font-size: 12px; border: 1px solid #00668a; font-family: Arial;" cellpadding="4px">
<tr style="background-color: #00668A; color: #FFFFFF; font-weight: bold; height: 20px;">
<td width="50%">&nbsp;Item</td>
<td width="50%">&nbsp;Result</td>
</tr>
"@


  foreach($item in (($InformationObject | Get-Member -MemberType Properties).Name)){
    $Report+=
    
    @"
<tr>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$item</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$(($InformationObject).$item)</td>
</tr>
"@
  }

  $Report+=
  @"
</table>
<br>
"@

  return $Report

}


#Little Helper
function Create-RegObject{

  param($BaseKey,
  [string[]] $RegValues)

  $result = New-Object psobject

  foreach($item in $RegValues){
    $AppVValue = (Get-ItemProperty -Path $BaseKey).$item
    $result | Add-Member -MemberType NoteProperty -Name $item -Value $AppVValue
  }

  return $result[0]

}

#Little Helper
function Get-MarkRedhtml{
  param($Keyword)
  if($keyword -eq  "NotDefined"){
    return '<font color="red">NotDefined<font color="black">'
  }

  return $keyword
}

#Little Helper
function Get-Highlightedhtml{
  param($Keyword)
  return $('<Mark>' + $keyword + '</Mark>')
}

#Little Helper
function Get-HighlightedIfTruehtml{
  param($Keyword)
  if($keyword -eq $true){
    return $('<Mark>' + $keyword + '</Mark>')
  }
  return $Keyword
}


#Little Helper
#Compare the field of a row and mark not idential
function Test-MarkItem{
  param(
    [Parameter(Mandatory=$false)] $element=$null,
    [Parameter(Mandatory=$true)][PSCustomobject []] $row,
    [Parameter(Mandatory=$false)][String] $configitem = $null
    
  ) 
 
  
  if(($null -eq $element) -or ($element -eq "") -or ($element -Match "NotDefined")){ #NotDefined is from the Module!
    return $false
  }
  
  #if($configitem -eq 'InProcessEnabled'){
  #  Write-Host Test
  #}
  
  
  foreach($item in $row){
    $notEual = 0
    $2Comparte = $item.$configitem
    
    #Empty List

    if(($null -ne$2Comparte) -and ($2Comparte.getType().name -eq "ArrayList")  -and ($2Comparte.count -eq 0))
    {
      $2Comparte = '<font color="red">NotDefined<font color="black">'
    }
      
    #Empty fiels
    if(($null -eq $2Comparte) -or ($2Comparte.ToString() -eq "") -or ($2Comparte.ToString() -eq "NotDefined")){
      $2Comparte = '<font color="red">NotDefined<font color="black">'
    }
      
   
    if(($null -ne $item.$configitem) -and ($item.$configitem.getType().name -eq "ArrayList")){
      $2Comparte = ($item.$configitem | Format-List | Out-String -Width 100) -replace("`n`r",'') -replace("`r",'<br>') -replace("^<br>",'')  -replace("<br>$",'') 
        
    }
    
    if(($2Comparte.ToString() -ne $element.toString()) -and (-not($2Comparte -match 'NotDefined'))){
      $notEual++
    }
  }
  if($notEual -ge 1){
    return $true
  }
  
  return $false
}

#Create a Comparsion betwenn configs
function Add-DeploymentComparsionReportHtml{
  param(
    [string] $Titel="Title",
    [string] $Commend ="Commend",
    [Parameter(Mandatory=$true)][string[]] $rows,
    [Parameter(Mandatory=$true)][String[]] $Columns,
    [Parameter(Mandatory=$true)][pscustomobject []] $Configurations
  )
    
  $Report = @"
<br><br>
<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">$Titel</span>
<br><br>
$Commend
<br><br>

<table width="95%" style="border-collapse: collapse; font-size: 12px; border: 1px solid #00668a; font-family: Arial;" cellpadding="4px">
<tr style="background-color: #00668A; color: #FFFFFF; font-weight: bold; height: 20px;">
<td>&nbsp;Setting</td>
"@

  foreach($c in $Columns){
    $Report += "<th>&nbsp;$c</th>"
  }
  $Report +='</tr>'

  
  foreach($item in $rows){
    #One Line first the property name
    $Report += "<tr><td style=`"border-right: 1px solid `#00668a; border-bottom: 1px solid `#00668a; `">$item</td>"
     
     
    foreach($conf in $Configurations){
      $cell = $conf.$item
      
      #Empty List
      if(($null -ne $cell) -and ($cell.getType().name -eq "ArrayList")  -and ($cell.count -eq 0)){
        $cell = '<font color="red">NotDefined<font color="black">'
      }
      
      #Empty fiels
      if(($null -eq $cell) -or ($cell.ToString() -eq "") -or ($cell.ToString() -eq "NotDefined")){
        $cell = '<font color="red">NotDefined<font color="black">'
        
      }
  
      

       
      if(($null -ne $cell) -and ($cell.getType().name -eq "ArrayList")){
        $cell = ($cell | Format-List | Out-String -Width 100) -replace("`n`r",'') -replace("`r",'<br>') -replace("^<br>",'')  -replace("<br>$",'') 
      }
       
              
      #Mark different Cells
      if($cell -match "NotDefined"){
        $mark = $false
      } else {
        $mark = Test-MarkItem -element $cell -row $Configurations -configitem $item  
      }       

       
      #Mark different Cells
      if(-not $mark){
        $Report += "<td style=`"border-right: 1px solid `#00668a; border-bottom: 1px solid `#00668a; `">$($cell)</td>"
      } else {
        $Report += "<td style=`"border-right: 1px solid `#00668a; border-bottom: 1px solid `#00668a; `"><mark>$($cell)</mark></td>"
      }
    }
    $Report += "</tr>"
  }

  $Report+=
  @"
</table>
<br>
"@

  return $Report
}


function Get-AppVServerReport{
  [CmdletBinding()]
  param()
  
  $IconSize = 14
  $AppVServerReport = $null
  
  $Date = Get-Date

  write-host
  write-host "------------------------------------------------------------------" -f yellow
  write-host "App-V 5.x Server Documentation Script"                              -f yellow
  write-host "------------------------------------------------------------------" -f yellow
  write-host "Script Started: $Date"                                  
  write-host


  $MachineName = $env:COMPUTERNAME
  $OperatingSystem = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name ProductName).ProductName

  write-host "Get server packages"  -f yellow
  
  #
  #
  # 4 Test
  #
  #
  #$Packagelist = Test-AppVServerPackageges
  $Packagelist = $Global:Packagelist
  #
  #
  #
  #
  #
  
  $AppVPNGImage = "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"

  $AppVServerReport = @"

<html>
<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<title>App-V 5.x Server Report</title>
</head>

<body style="font-family: Arial; font-size: 12px;">
<!-- IMAGEHERE -->
<span style="font-size: 38px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">App-V 5.x Server Report</span>
<br><br><br>
(c) Andreas Nick 2020<br>
Free for use at your own risk (MIT License)<br>
<a href="https://www.AndreasNick.com" target="_blank">www.AndreasNick.com</a><br>
<a href="https://www.software-virtualisierung.de" target="_blank">www.software-virtualisierung.de</a><br>
<a href="https://www.nick-it.de" target="_blank">www.nick-it.de</a><br>

<br>
This document provides information on the current status of the App-V 5.x Server, the following results are below:
<br><br>
<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">Machine Information</span>
<br><br>

<table width="30%" style="border-collapse: collapse; font-size: 12px; border: 1px solid #00668a; font-family: Arial;" cellpadding="4px">
<tr style="background-color: #00668A; color: #FFFFFF; font-weight: bold; height: 20px;">
<td width="50%">&nbsp;Item</td>
<td width="50%">&nbsp;Result</td>
</tr>
<tr>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">Machine Name</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$MachineName</td>
</tr>
<tr>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">Operating System</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$OperatingSystem</td>
</tr>
<tr>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">Date</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$Date</td>
</tr>
</table>
<br>

<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">App-V Client Information</span>
<br><br>
<span style="color:green;">&#x2714;</span> - Enabled / Running / Passed
<br>
<span style="color:red;">&#x2718;</span> - Disabled / Stopped / Failed
<!-- br>
<span style="color:navy">&equiv;</span> - Set via Group Policy
<br>
<span style="color:orange">&ne;</span> - Set Locally
<br -->
<br>
"@

  $AppVServerReport = $AppVServerReport -replace '<!-- IMAGEHERE -->', $('<img src="data:image/png;base64,' + $AppVPNGImage + '" width="64" height="64"  title="'+"https://www.software-virtualisierung.de" +'" />')


  write-host "Get Management server data"                              -f yellow


  $Vals = Create-RegObject -BaseKey "HKLM:SOFTWARE\Microsoft\AppV\Server\ManagementService" -RegValues @("MANAGEMENT_WEBSITE_NAME","MANAGEMENT_DB_SQL_SERVER","MANAGEMENT_SQL_CONNECTION_STRING","MANAGEMENT_ADMINACCOUNT_SID","MANAGEMENT_CONSOLE_URL")
  $PublishingServerReport = $(Add-ReportHtml -InformationObject $Vals -Titel "Management Server Report" -Commend "Data from the App-V management server" )
  $AppVServerReport+= $PublishingServerReport

  write-host "Get Publishing Server Data"                              -f yellow
  
  $Vals = Create-RegObject -BaseKey "HKLM:SOFTWARE\Microsoft\AppV\Server\PublishingService" -RegValues @("PUBLISHING_WEBSITE_NAME","PUBLISHING_MGT_SERVER_TIMEOUT","PUBLISHING_MGT_SERVER_REFRESH_INTERVAL","PUBLISHING_WEBSITE_PORT","PUBLISHING_MGT_SERVER")
  $PublishingServerReport = $(Add-ReportHtml -InformationObject $Vals -Titel "Publishing Server Report" -Commend "Data from the App-V Publishing server" )
  $AppVServerReport+= $PublishingServerReport
  

  ########################################
  # Report - Test-AppVPackagePath
  ########################################

  Write-host "Creating Report Data for $MachineName - Package Path Status..." -ForegroundColor Yellow

  $AppVServerReport+= @"

<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">Server packages</span>
<br><!-- br>
This check validates whether the Package Path is still valid, if the path isn't valid and your running in Shared Content Store mode the application won't launch.
<br>
<br>
This script will return the following results:
<br><br>
<span style="color:green;">&#x2714;</span> - Package Path Valid
<br>
<span style="color:red;">&#x2718;</span> - Package Path can't be found
<br --><br>
<table width="95%" style="border-collapse: collapse; font-size: 12px; border: 1px solid #00668a; font-family: Arial;" cellpadding="4px">
<tr style="background-color: #00668A; color: #FFFFFF; font-weight: bold; height: 20px;">
<td width="1%">&nbsp;Count</td>
<td width="1%">&nbsp;ID</td>
<td width="15%">&nbsp;Name</td>
<td width="5%">&nbsp;Version</td>
<td width="10%">&nbsp;PackageID</td>
<td width="10%">&nbsp;VersionID</td>
<td width="20%">&nbsp;PackagePath</td>
<td width="5%" align="center">&nbsp;Valid Path</td>
<td width="5%" align="center">&nbsp;Enabled</td>
</tr>

"@


  ###########################################
  #Management
  ###########################################

  Write-host "Collating Data for the App-V Database" -ForegroundColor Yellow
  #$AppVClientVersion = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\AppV\Client -Name Version).Version


  foreach($Pack in $Packagelist){

    $AppVServerReport+= @"
<tr>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$($Pack.count)</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$($Pack.ID)</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$($Pack.PackageName)</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$($Pack.PackageVersion)</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$($Pack.PackageID)</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$($Pack.VersionID)</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$($Pack.PackagePath)</td>
"@

    if($Pack.PackageTestPath -eq "True"){

      $AppVServerReport+= @"
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; color:green; font-size:$IconSize;" align="center">&#x2714;</td>
"@

    }

    elseif($Pack.PackageTestPath -eq "False"){

      $AppVServerReport+= @"
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; color:red; font-size:$IconSize;" align="center"> &#x2718;</td>
"@

    }

    if($Pack.Enabled -eq "True"){
      $AppVServerReport+= @"
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; color:green; font-size:$IconSize;" align="center">&#x2714;</td>
"@
    } else {
      $AppVServerReport+= @"
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; color:red; font-size:$IconSize;" align="center"> &#x2718;</td>
"@
    }
  }

  $AppVServerReport+= @"
</table>
<br>
<br>
"@

  ########################################
  # Report - Package Analysis
  ########################################

  Write-host "Creating Report Data for $MachineName - Package Path Status..." -ForegroundColor Yellow


  $AppVServerReport+= @"
<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">App-V Server Packages Detail Info</span>
<br>
Get Infos from the packages: Com Mode, Services, Global Objects, AD Groups, ConnectionGroups.
<br>
<br>
<table width="95%" style="border-collapse: collapse; font-size: 12px; border: 1px solid #00668a; font-family: Arial;" cellpadding="4px">
<tr style="background-color: #00668A; color: #FFFFFF; font-weight: bold; height: 20px;">
<td width="5">&nbsp;On</td>
<td >&nbsp;Name</td>
<td  align="Center">&nbsp;AD-Groups</td>
<td  align="Center">&nbsp;Connection-Groups</td>
<td  align="Center">&nbsp;Settings intern<br>(from AppXMaifest)</td>
<td  align="Center">&nbsp;Settings server<br>(Get-AppVSrvPackageDeplCfg)</td>
</tr>
"@

  foreach($Pack in $Packagelist){

    $ADGroups =""

    if(@($Pack.Entitlements).count -gt 0){
      $ADGroups = [string]::Join("</br>", $($Pack.Entitlements))
    }
    $ConGroups =""
    if($Pack.ConnectionGroups.count -gt 0){
      $ConGroups = [string]::Join("</br>",$($Pack.ConnectionGroups.Name))
    }
    
    $AppVServerReport+="<tr>"

    if($Pack.Enabled -eq "True"){
      $AppVServerReport+= @"
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; color:green; font-size:$IconSize;" align="center">&#x2714;</td>
"@

    } else {
      $AppVServerReport+= @"
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; color:red; font-size:$IconSize;" align="center"> &#x2718;</td>
"@

    }    
    #<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$($Pack.ID)</td>
    $htmlBlock = @"


<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;"><h3>$($Pack.PackageName)</h3><!-- ICONSHERE --></td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$ADGroups</td>
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a;">$ConGroups</td>
"@
    #AddIcons
    
    if($ExtractIcon){
      if($pack.Manifestinfo.HasShortcuts -and (@($pack.Icons).count -gt 0)) {
        $Imageblock = "<br>"
        foreach($icon in @($pack.Icons)) {
          $Imageblock +='<table border="0"  style="font-size: 12px; font-family: Arial;" cellpadding="4px"><tr><td>'
          $Imageblock += $("`r`n" + '<img src="data:image/png;base64,' + $icon.Base64Image + '" width="32" height="32"  title="'+$icon.file + "`n" + $icon.Icon + '" />')
          $Imageblock += '</td><td>' + $(Split-Path -leaf $icon.Target)+ ' '+ $icon.Arguments + "<br>" + ([string]$icon.file).Substring(0,([string]$icon.file).IndexOf(']')+1) + '</td></tr></table>'
        }
        $htmlBlock  = $htmlBlock -replace '<!-- ICONSHERE -->',$Imageblock 
      
 
      }
    }
    
    $AppVServerReport+=$htmlBlock

    ###########################################
    # Package intern settings
    ###########################################
    

    $ComModeString = $($pack.ManifestInfo.ComMode)
    
    if($ComModeString -eq "Integrated")
    {
      $ComModeString = Get-Highlightedhtml $ComModeString
    }
    
    $ObjModeString = $($pack.ManifestInfo.ObjectsMode)
    
    if($ObjModeString -eq "NotIsolate")
    {
      $ObjModeString = Get-Highlightedhtml $ObjModeString
    }
    
    
    $AppVServerReport+= @"
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; ">Mode=$ComModeString</br>
OutOfProcessEnabled=$($pack.ManifestInfo.OutOfProcessEnabled)</br>
InProcessEnabled=$($pack.ManifestInfo.InProcessEnabled)</br>
GlobalObjects=$($ObjModeString)</br>
FullVFSWriteMode=$(Get-HighlightedIfTruehtml($pack.ManifestInfo.FullVFSWriteMode))</br>
HasBrowserHelpObjects=$(Get-HighlightedIfTruehtml($pack.ManifestInfo.HasBrowserHelpObject))</br>
HasEnvironmentVariables=$(Get-HighlightedIfTruehtml($pack.ManifestInfo.HasEnvironmentVariables))<br>
HasHasFonts=$(Get-HighlightedIfTruehtml($pack.ManifestInfo.HasFonts))<br>
HasServices=$(Get-HighlightedIfTruehtml($pack.ManifestInfo.HasServices))<br>
HasShellExtensions=$(Get-HighlightedIfTruehtml($pack.ManifestInfo.HasShellExtensions))<br>
HasUserScripts=$(Get-HighlightedIfTruehtml($pack.ManifestInfo.HasUserScripts))<br>
HasMachineScripts=$(Get-HighlightedIfTruehtml($pack.ManifestInfo.HasMachineScripts))
</td>
"@

    ###########################################
    # Package Server settings
    ###########################################

    $ComModeString = $($pack.DeploymentConfigInfo.ComMode)
    
    if($ComModeString -eq "Integrated")
    {
      $ComModeString = Get-Highlightedhtml $ComModeString
    }
    
    $ObjModeString = $($pack.DeploymentConfigInfo.ObjectsEnabled)
    
    if($ObjModeString -eq $false)
    {
      $ObjModeString =  Get-Highlightedhtml $ObjModeString 
    }
    
    
    $AppVServerReport+= @"
<td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; ">Mode=$(Get-MarkRedhtml($ComModeString))</br>
OutOfProcessEnabled=$(Get-MarkRedhtml($pack.DeploymentConfigInfo.OutOfProcessEnabled))</br>
InProcessEnabled=$(Get-MarkRedhtml($pack.DeploymentConfigInfo.InProcessEnabled))</br>
ObjectsEnabled=$(Get-MarkRedhtml($ObjModeString))</br>
FullVFSWriteMode=<font color="red">NotDefined<font color="black"></br>
HasBrowserHelpObjects=<font color="red">NotDefined<font color="black"></br>
HasEnvironmentVariables=$(Get-MarkRedhtml($pack.DeploymentConfigInfo.HasEnvironmentVariables))</br>
HasHasFonts=$(Get-MarkRedhtml($pack.DeploymentConfigInfo.FontsEnabled))</br>
HasServices=$(Get-MarkRedhtml($pack.DeploymentConfigInfo.ServicesEnabled))</br>
HasURLProtocols=$(Get-HighlightedIfTruehtml($pack.DeploymentConfigInfo.HasURLProtocols))</br>
HasUserScripts=$(Get-HighlightedIfTruehtml($pack.DeploymentConfigInfo.HasUserScripts))</br>
HasMachineScripts=$(Get-HighlightedIfTruehtml($pack.DeploymentConfigInfo.HasMachineScripts))

</td>
"@

  }

  $AppVServerReport+= @"
</table>
<br>
<br>
"@


  #
  #
  # Compare Package settings in the package, on disk, and on the Server. 
  #
  #
   
  if($FullReport){ 
   
    Write-host "Creating full Package report..." -ForegroundColor Yellow
    
    $AppVServerReport += '<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">App-V Package Settings Comparsion</span><br>'
    $AppVServerReport += 'we compare the setting in the package with the settings on the server and the DeploymentConfig.xml<br>'
    $AppVServerReport += '(with the same name) on the hard disk. You can see where the scripts came from or if maybe an import was forgotten.<br>'
    $AppVServerReport += 'Furthermore, detailed information about the script triggers is provided here'

  
    #Elements to Compare
    $InformationElements = @("Name, ApplicationCapabilities, ApplicationCapabilitiesEnabled, Applications, ComMode, OutOfProcessEnabled, InProcessEnabled, EnvironmentVariablesEnabled, FileSystemEnabled, FileTypeAssociationEnabled, FileTypeAssociation,
        FontsEnabled, HasApplicationCapabilities, HasApplications, HasEnvironmentVariables, HasFileTypeAssociation, HasMachineRegistrySettings, HasRegistrySettings,
        HasShortcuts, Shortcuts, HasTerminateChildProcesses, HasURLProtocols, HasUserScripts, UserScripts, HasMachineScripts, MachineScripts, ObjectsEnabled, RegistryEnabled, ServicesEnabled, ShortcutsEnabled,
    TerminateChildProcesses, HasFonts, HasServices, Services, HasShellExtensions, HasSxSAssemblys, SxSAssemblys, PackageFullLoad, FullVFSWriteMode,  ObjectsMode".replace("`n", "").replace("`r", "").replace(" ", "").split(','))
  
  
    foreach($Pack in $Packagelist){
      Write-Verbose "Create comparsion for package $($pack.ManifestInfo.Name)"
      $Minfo = $Pack.ManifestInfo
      $DeplInfo = $Pack.DeploymentConfigInfo
      $DepFromDisk = $Pack.Disk_DeploymentConfigInfo
      if($null -eq $DepFromDisk){
        $DepFromDisk = new-Object PSCustomObject
      }
     
      $AppVServerReport += Add-DeploymentComparsionReportHtml -Titel "Configuration Comparison $( $pack.ManifestInfo.Name)  $($pack.PackagePath)  Version: $($pack.PackageVersion) " -Commend "Note: Not all items are included in all configuration files!" -rows $InformationElements `
      -Columns @("From Manifest", "From Server", "Config From Disk") `
      -Configurations @($Minfo, $DeplInfo, $DepFromDisk)

    }
  }
  
  
  
  #
  #
  # Report Packages with ShellExtensions
  #
  #

  $AppVServerReport += '<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">App-V Packages With ShellExtensions</span><br>'
  $AppVServerReport+= @"
<br>
<table width="60%" style="border-collapse: collapse; font-size: 12px; border: 1px solid #00668a; font-family: Arial;" cellpadding="4px">
<tr style="background-color: #00668A; color: #FFFFFF; font-weight: bold; height: 20px;">
<td >&nbsp;Name</td>
<td  align="Center">&nbsp;HasShellExtensions</td></tr>
"@
  $hasShellExtensions = $Packagelist | ForEach-Object {$_.ManifestInfo | Where-Object {$_.HasShellExtensions -ne $false} | Select-Object -Property Name, HasShellExtensions} 
  $AppVServerReport+= $hasShellExtensions | ForEach-Object {
     
    @"
    <tr><td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; ">$($_.Name)</td>
    <td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; ">$($_.HasShellExtensions)</td></tr>
"@
  }

  $AppVServerReport+= "</table><br><br>"

  #
  #
  # Report Packages with SXSAssambies
  #
  #
  
  $AppVServerReport += '<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">App-V Packages With Assemblies</span><br>'

  $AppVServerReport+= @"
<br>
<table width="60%" style="border-collapse: collapse; font-size: 12px; border: 1px solid #00668a; font-family: Arial;" cellpadding="4px">
<tr style="background-color: #00668A; color: #FFFFFF; font-weight: bold; height: 20px;">
<td >&nbsp;Name</td>
<td  align="Center">&nbsp;SXSAssemblies</td></tr>
"@
  $SXSAssemblies = $Packagelist | ForEach-Object {$_.ManifestInfo | Where-Object {$_.HasSxSAssemblys -ne $false} | Select-Object -Property Name, SxSAssemblys} 
  $AppVServerReport+= $SXSAssemblies | ForEach-Object {
     
    @"
    <tr><td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; ">$($_.Name)</td>
    <td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; ">$($_.SxSAssemblys)</td></tr>
"@
  }

  $AppVServerReport+= "</table><br><br>"
  
  
  #
  #
  # Report Packages with Services
  #
  #
  
  $AppVServerReport += '<span style="font-size: 14px; font-weight: bold; color: #006699; text-decoration: underline; font-family: Arial;">App-V Packages With Services</span><br>'

  $AppVServerReport+= @"
<br>
<table width="60%" style="border-collapse: collapse; font-size: 12px; border: 1px solid #00668a; font-family: Arial;" cellpadding="4px">
<tr style="background-color: #00668A; color: #FFFFFF; font-weight: bold; height: 20px;">
<td >&nbsp;Name</td>
<td  align="Center">&nbsp;Services</td></tr>
"@
  $Services = $Packagelist | ForEach-Object {$_.ManifestInfo | Where-Object {$_.HasServices -ne $false} | Select-Object -Property Name, Services} 
  $AppVServerReport+= $Services | ForEach-Object {
     
    @"
    <tr><td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; ">$($_.Name)</td>
    <td style="border-right: 1px solid #00668a; border-bottom: 1px solid #00668a; ">$(($_.Services | Out-String -Width 100)  -replace("`n`r",'') -replace("`r",'<br>') -replace("^<br>",'')  -replace("<br>$",''))</td></tr>
"@
  }

  $AppVServerReport+= "</table><br><br>"
    
  write-host
  write-host "Create report finished: $Date"                                  
  return $AppVServerReport

}

$Global:Packagelist = Test-AppVServerPackageges

$AppVServerReport = Get-AppVServerReport 

$AppVServerReport > $OutputPath

Write-Host "Created the Report in: file://$OutputPath" -ForegroundColor Cyan
Write-Host "please try the script also with buttons for a full report with icons `nTest-AppVServer.ps1 -StartBrowser -FullReport -ExtractIcon" -ForegroundColor Cyan

Write-Host "Please visit my blogs www.AndreasNick.com or www.Software-Virtualisierung.de`nyou can follow me on Twitter under @Nickinformation" -ForegroundColor Green -BackgroundColor Blue

if($StartBrowser){
  Invoke-Item $OutputPath
} else {
  return  $OutputPath
}


