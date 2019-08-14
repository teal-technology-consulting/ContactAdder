#Variables
[bool]$initialsetup = $false
$latestcontactadderscript = Get-ChildItem $PSScriptroot -filter "public_ContactAdder*.ps1" | sort LastWriteTime | select -last 1
[string]$scheduledTaskFolder = "PublicTeal"
[string]$ScheduledTaskName = "PublicTeal_Contact_Adder"
[string]$ScheduledTaskDescription = "Scheduled Task to run the Teal Contact Adder to update and create new Teal Contact from the O365 GAL."
[string]$ScheduledTaskExecutable = "powershell.exe"
[string]$ScheduledTaskWorkDir = $PSScriptroot
[string]$ScheduledTaskArguments = "-ExecutionPolicy Bypass -windowstyle Hidden "+$latestcontactadderscript.FullName +"; exit $"+"LASTEXITCODE"
[string]$ScheduledTaskState = "Ready"
[int]$scheduledTaskRuntime = "5"
[array]$ScheduledTaskExists = @()
$ScheduledTaskUser = "$($env:USERDOMAIN)\$($env:USERNAME)"

#Define Logging:
#Logging
$LogfileBasePath = $PSScriptRoot + "\log\"
$LogfileDate = (Get-Date).ToString("ddMMyyyy-HHmm")
$logname = $(($MyInvocation.MyCommand.Name).Split(".")[0])
$logfile = $LogfileBasePath + $logname +"_"+$LogfileDate+".log"
[int]$keeplatestlogcount = 5



#region functions

function Write-Log {
     [CmdletBinding()]
     param(
         [string]$LogFile,
         [string]$Classification,
         [string]$Message
     )

    $message = "$((Get-Date).ToString("ddMMyyyy-HHmmssffff"))`t$($Classification.ToUpper())`t$Message"
    [System.IO.File]::AppendAllText($LogFile, "$message`r`n",[System.Text.UTF8Encoding]::UTF8)
    }

function delete-oldlogs {
    param (
        [int]$keeplatestlogcount,
        [string]$logpath,
        [string]$logname
        )
    [array]$keeplogs = $null
    $keeplogs = (Get-ChildItem $logpath -filter "$logname*.log" | sort LastWriteTime | select -last $keeplatestlogcount).fullname
    foreach ($logtodelete in gci $logpath -filter "$logname*.log")
        {
        if ($keeplogs -notcontains $logtodelete.fullname){
            Remove-Item $logtodelete.fullname
            }
        }
    }

function New-TaskFolder
    {
    Param ($TaskPath)
    if (!($TaskPath)){return}
    $ErrorActionPreference = "stop"
    $ScheduleObject = New-Object -ComObject schedule.service
    $ScheduleObject.connect()
    $RootFolder = $ScheduleObject.getfolder("\")
        Try {$null = $ScheduleObject.getfolder($TaskPath)}
        catch {$null = $RootFolder.createfolder($TaskPath)}
        Finally { $ErrorActionPreference = "continue"}
    }

function get-Task {
    param (
    $TaskPath,
    $TaskName)
    $taskexists = $null
    try {
        $TaskExists = Get-ScheduledTask -TaskName $TaskName -TaskPath $("\$TaskPath\") -ErrorAction Stop
        return $TaskExists
        }
    catch{
        return $TaskExists
        }
    
    }

function New-Task
    {
    param(
    $TaskName,
    $TaskWorkDir,
    $TaskPrincipal,
    $TaskDescription,
    $TaskPath,
    $TaskExecutable,
    $TaskArgument,
    $TaskPassword,
    $TaskRuntime)

    $TaskAction=New-ScheduledTaskAction -Execute "$TaskExecutable" -WorkingDirectory $TaskWorkDir -Argument "$TaskArgument"
    $TaskSettings=New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -Compatibility Win8 -Priority 7 -ExecutionTimeLimit (New-TimeSpan -Minutes $TaskRuntime) -StartWhenAvailable
    $TaskTrigger=New-ScheduledTaskTrigger -Daily -At 9am
    $Task=Register-ScheduledTask "$TaskName" -Action $TaskAction -Description $TaskDescription -User $TaskPrincipal -Settings $TaskSettings -TaskPath $TaskPath -Trigger $TaskTrigger
    return $SchTask.State
    }

function Update-Task {
    param(
    $TaskName,
    $TaskWorkDir,
    $TaskPrincipal,
    $TaskDescription,
    $TaskPath,
    $TaskExecutable,
    $TaskArgument,
    $TaskPassword,
    $TaskRuntime)
    $TaskTrigger=New-ScheduledTaskTrigger -Daily -At 9am
    $TaskAction=New-ScheduledTaskAction -Execute "$TaskExecutable" -WorkingDirectory $TaskWorkDir -Argument "$TaskArgument"
    $TaskSettings=New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -Compatibility Win8 -Priority 7 -ExecutionTimeLimit (New-TimeSpan -Minutes $TaskRuntime) -StartWhenAvailable
    $Task=Set-ScheduledTask -TaskName $TaskName -TaskPath $("\$TaskPath\") -Action $TaskAction -Settings $TaskSettings -Trigger $TaskTrigger
    return $Task.state
    } 

#endregion functions

if (!(test-path $LogfileBasePath)){
    New-Item -Path $LogfileBasePath -ItemType Directory -Force
    }

#region remove old logs

delete-oldlogs -keeplatestlogcount $keeplatestlogcount -logpath $LogfileBasePath -logname $logname

#endregion remove old logs

Write-Log -LogFile $logfile -Classification "INFO" -Message "Script start ..."

#region create Scheduled TaskFolder
Write-Log -LogFile $logfile -Classification "info" -Message "Create Scheduled Task folder .... $scheduledTaskFolder"
try {
    New-TaskFolder -taskpath $scheduledTaskFolder
    Write-Log -LogFile $logfile -Classification "Success" -Message "Scheduled Task folder $scheduledTaskFolder created"
    }

catch{
    Write-Log -LogFile $logfile -Classification "error" -Message "cheduled Task folder $scheduledTaskFolder not created: $($error[0])"
    }
#endregion create Scheduled TaskFolder

#region Check Task exists
Write-Log -LogFile $logfile -Classification "info" -Message "Check if the tasks exists ..."
try {
    $ScheduledTaskExists = get-Task -TaskName $ScheduledTaskName -TaskPath $scheduledTaskFolder
    Write-Log -LogFile $logfile -Classification "success" -Message "Scheduled Task exists." 
    }
catch {
    Write-Log -LogFile $logfile -Classification "success" -Message "Scheduled Task dont exists." 
    }

#endregion Check Task exists

#region Create Task
Write-Log -LogFile $logfile -Classification "info" -Message "Start creating or updating the scheduled task ..."
try {
    if (($ScheduledTaskExists).state -like "Ready") {
        Update-Task -TaskName $ScheduledTaskName -TaskPath $scheduledTaskFolder -TaskExecutable $ScheduledTaskExecutable -TaskWorkDir $ScheduledTaskWorkDir -TaskArgument $ScheduledTaskArguments -TaskRuntime $scheduledTaskRuntime
        Write-Log -LogFile $logfile -Classification "success" -Message "Scheduled Task updated."
        }

    if (($ScheduledTaskExists).state -like "Running") {
        Write-Log -LogFile $logfile -Classification "error" -Message "Scheduled Task is running update not possible."
        throw "Scheduled Task is running update not possible."
        }

    if (!($ScheduledTaskExists)){
        New-Task -TaskName $ScheduledTaskName -TaskDescription $ScheduledTaskDescription -TaskWorkDir $ScheduledTaskWorkDir -TaskExecutable $ScheduledTaskExecutable -TaskPath $scheduledTaskFolder -TaskArgument $ScheduledTaskArguments -TaskPrincipal $ScheduledTaskUser -TaskRuntime $scheduledTaskRuntime
        Write-Log -LogFile $logfile -Classification "info" -Message "Scheduled Task created."
        $initialsetup = $true
        }
    }
catch{
    Write-Log -LogFile $logfile -Classification "error" -Message "There was a problem with the Scheduled Task setup: $($error[0]) please fix the error and start the script again."
    }

#region Create Task

#region initale run of contactadder to create the contacts.
if ($initialsetup){
    try{
        Write-Log -LogFile $logfile -Classification "Info" -Message "Run Scheduled task for initial creation of contacts..."
        Start-ScheduledTask -TaskName $ScheduledTaskName -TaskPath "\$scheduledTaskFolder"
        Write-Log -LogFile $logfile -Classification "Success" -Message "Run Scheduled task for initial creation of contacts..."
        }
    catch {
        Write-Log -LogFile $logfile -Classification "Error" -Message "There was a problem to execute the scheduled task."
        Write-Log -LogFile $logfile -Classification "Error" -Message "$($Eror[0])"
        exit 99
        }
    }
#endregion