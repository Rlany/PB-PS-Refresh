Param(
    $sourceWorkbook = "Demo.pbix",
    $workbook = "Demo_New.pbix",
    $closeWB = $true
)

$ErrorActionPreference = 'Stop'

# add reference to automation functions
. .\PBUIAutomation.ps1

##################################################################################
##        BODY
##################################################################################


$folder = $PSScriptRoot 
$srcFile = $folder + "\" + $sourceWorkbook
$tgtFile = $folder + "\" + $workbook

#kill process if the workbook is opened
#kill process if the workbook is opened and start new one
#referenced from PBUIAutomation.ps1
KillOldPB -targetWorkbook $tgtFile;


#copy workbook under new name
Copy-Item -Path $srcFile -Destination $tgtFile -force


$workbookName = (get-item $tgtFile).BaseName
$appWindowName = "$workbookName - Power BI Desktop"
$refreshButtonName = "Refresh"
$refreshWindowName = "Refresh"

#Open Power BI workbook
Invoke-Item $tgtFile
 
#Wait for main window of Power BI Desktop  
#referenced from PBUIAutomation.ps1 
WaitForPBtoOpen -appWindowName $appWindowName

# get app ID so we can check after refresh is hit running modal window with title refresh
$appUI = FindProcessUI -appName $appWindowName;


#Press REFRESH
FindAndClickByName -name $refreshButtonName -appUI $appUI
"Refreshing"

Start-Sleep -Seconds 1

# find refreshing window and wait until it is running
$refreshOpen = $true
while ($refreshOpen) {
    try{
        $refreshWindow = FindModalWindow -name $refreshWindowName -appUIObject $appUI
        $refreshOpen = ($null -ne $refreshWindow);
        #if refresh window is open then wait for 10 seconds and try again
        if ($refreshOpen) {
            Start-Sleep -Seconds 10
            "*"
        }
    }
    catch{
        $refreshOpen = $false;
    }
}


#Refreshing windows is not running anymore
"Data refreshed locally"


# save
FindAndClickByName -name "Save" -appUI $appUI
"File successfully saved"

Start-Sleep -Seconds 3

#close
if ($closeWB) 
{
    FindAndClickByName -name "Close" -appUI $appUI
    "Successfully closed"
}
<#
this is code for playing around
$cond5 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, "System Menu Bar")
$menu = $appUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $cond2)

$cond4 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::IsEnabledProperty, $true)
$appPartSaveAS = $menu.FindAll([Windows.Automation.TreeScope]::Descendants, $cond4)

foreach($qq in $appPartSaveAS)
{ 
    $cn = $qq.Current.ClassName;
    write-host $cn
    $cc = $qq.Current;
    write-host $cc;
    write-host $cc.Name;
    if($cc.Name -like "*menu*")
    {
        write-host $cc.Name;    
    }
    $ccs = $cc;
}
return;


#>