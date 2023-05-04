
$ErrorActionPreference = 'Stop'

Add-Type –AssemblyName UIAutomationClient
Add-Type –AssemblyName UIAutomationTypes

##################################################################################
##        Functions
##################################################################################

#------------------------------------------------------------------------------------------------
# Find PowerBI Main app UI ID
function FindProcessUI([string]$appName)
{
    $processId = ((Get-Process).where{$_.MainWindowTitle -eq $appName})[0].Id
    $root = [Windows.Automation.AutomationElement]::RootElement
    $condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $processId)
    $appUI = $root.FindFirst([Windows.Automation.TreeScope]::Children, $condition)
    return $appUI
}

#------------------------------------------------------------------------------------------------
# Find componet of application by class name and name of element
function FindAppPart([string]$name, [string]$className, [System.Windows.Automation.AutomationElement]$appUI, [boolean]$useClassName = $true)
{
    $cond = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $name)
    if($useClassName)
    {
        $cond2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ClassNameProperty, $className)
        $cond3 = New-Object Windows.Automation.AndCondition($cond, $cond2)
    }
    else
    {
        $cond3 = $cond;
    }
    $appPart = $appUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $cond3)
   # $cn = $appPart.Current.ClassName;
   # write-host $cn
   # $cc = $appPart.Current;
   # write-host $cc
   # $ccs = $cc;
    return $appPart;
}

#------------------------------------------------------------------------------------------------
# Find button by name
function FindButton([string]$name, [System.Windows.Automation.AutomationElement]$appUI){
    
	$button = FindAppPart -name $name -className "ms-Button root-199" -appUI $appUI
    return $button 
}

#------------------------------------------------------------------------------------------------
# Find modal winbdow by name
function FindModalWindow([string]$name, $appUIObject){
    
	$window = FindAppPart -name $name -className "modal-frame" -appUI $appUIObject
    return $window
}

#------------------------------------------------------------------------------------------------
# Find button by name and invoke click
function FindAndClickButton([string]$name, [System.Windows.Automation.AutomationElement]$appUI){
	
    $button = FindButton -name $name -appUI $appUI
	$button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
}

#------------------------------------------------------------------------------------------------
# Find object by name
function FindByName([string]$name, [System.Windows.Automation.AutomationElement]$appUI){
    
	$button = FindAppPart -name $name -className "none" -appUI $appUI -useClassName $false
    return $button 
}


# Find object by name and invoke click
function FindAndClickByName([string]$name, [System.Windows.Automation.AutomationElement]$appUI){
	
    $button = FindByName -name $name -appUI $appUI
	$button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
}

# This method kills PB if it's running already with existing workbook
# and then runs PB with reference of the workbook
function KillOldPB([string]$targetWorkbook){
    
    $workbookName = (get-item $targetWorkbook).BaseName
    $appWindowName = "$workbookName - Power BI Desktop"
    $processId = ((Get-Process).where{$_.MainWindowTitle -eq $appWindowName})[0].Id
    
    if($processId ){
        "Power BI workbook $workbookName is already open. Killing process..."
        taskkill /F /PID $processId
        Start-Sleep -Seconds 2
    }
}

# this method will keep sleeping until power bi deskto is open.
# we identify this by finding app window and finding enabled refresh button.
function WaitForPBtoOpen([string]$appWindowName){
    $ready = $false;
    "Waiting for Power BI to open"
    #We are waiting until refresh button is ready and enabled - this mean mode is loaded
    while (!$ready){
        try{
            $xAppUI = FindProcessUI -appName $appWindowName
            $refreshButton = FindByName -name "Refresh" -appUI $xAppUI
            $ready = !($null -eq $refreshButton)
        }
        catch{
            Start-Sleep -Seconds 3
            "*"
        }
    }
    "Power BI opened"
}