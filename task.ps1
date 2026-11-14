param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Register','Test','Status','Enable','Disable','Remove')]
    [string]$Action
)

$ErrorActionPreference = 'Stop'
$Name = 'USDA WASDE Excel Auto Update'
$TaskPath = '\'
$Root = $PSScriptRoot
$Python = Join-Path $Root '.venv\Scripts\python.exe'
$Main = Join-Path $Root 'wasde_automation.py'
$Config = Join-Path $Root 'config.json'
$ArgsBase = '"' + $Main + '" --config "' + $Config + '"'

function Register-WasdeTask($TaskName, $Trigger, $Arguments) {
    $Command = New-ScheduledTaskAction -Execute $Python -Argument $Arguments -WorkingDirectory $Root
    $Identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $Principal = New-ScheduledTaskPrincipal -UserId $Identity -LogonType Interactive -RunLevel Limited
    $Settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
    Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $Command `
        -Principal $Principal -Trigger $Trigger -Settings $Settings | Out-Null
}

if ($Action -in 'Register','Test') {
    foreach ($Path in $Python,$Main,$Config) {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing $Path" }
    }
}

switch ($Action) {
    'Register' {
        if ((Get-TimeZone).Id -ne 'Turkey Standard Time') {
            throw 'Registration expects the Windows Istanbul time zone (Turkey Standard Time).'
        }
        & $Python $Main --config $Config --check-install
        if ($LASTEXITCODE -ne 0) { throw 'Installation check failed. See runtime\wasde_automation.log.' }
        $Now = Get-Date
        $Calendar = Get-Content -LiteralPath (Join-Path $Root 'release-dates.json') -Raw | ConvertFrom-Json
        $Triggers = @(foreach ($Entry in $Calendar.PSObject.Properties) {
            $Start = [datetime]::ParseExact($Entry.Value + ' 18:55', 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
            if ($Start.AddHours(3) -ge $Now) {
                New-ScheduledTaskTrigger -Once -At $Start `
                    -RepetitionInterval (New-TimeSpan -Minutes 2) -RepetitionDuration (New-TimeSpan -Hours 3)
            }
        })
        if ($Triggers.Count -eq 0) { throw 'No current or future dates in release-dates.json.' }
        Register-WasdeTask $Name $Triggers ($ArgsBase + ' --scheduled')
        Write-Host "Registered $Name with $($Triggers.Count) release dates."
    }
    'Test' {
        $TestName = $Name + ' - Test ' + [Guid]::NewGuid().ToString('N')
        $Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddDays(1)
        Register-WasdeTask $TestName $Trigger ($ArgsBase + ' --check-install')
        try {
            $Before = (Get-ScheduledTaskInfo -TaskName $TestName -TaskPath $TaskPath).LastRunTime
            Start-ScheduledTask -TaskName $TestName -TaskPath $TaskPath
            $Deadline = (Get-Date).AddMinutes(2)
            do {
                Start-Sleep -Seconds 1
                $Current = Get-ScheduledTask -TaskName $TestName -TaskPath $TaskPath
                $Info = $Current | Get-ScheduledTaskInfo
                if ($Info.LastRunTime -ne $Before -and $Current.State -eq 'Ready') {
                    if ($Info.LastTaskResult -ne 0) {
                        throw "Scheduled test failed (code $($Info.LastTaskResult)). See runtime\wasde_automation.log."
                    }
                    Write-Host 'PASS: Task Scheduler launched Python and Excel COM under the current account.'
                    return
                }
            } while ((Get-Date) -lt $Deadline)
            throw 'Scheduled test timed out.'
        }
        finally {
            Stop-ScheduledTask -TaskName $TestName -TaskPath $TaskPath -ErrorAction SilentlyContinue
            Unregister-ScheduledTask -TaskName $TestName -TaskPath $TaskPath -Confirm:$false
        }
    }
    'Status' { Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath | Get-ScheduledTaskInfo }
    'Enable' { Enable-ScheduledTask -TaskName $Name -TaskPath $TaskPath }
    'Disable' { Disable-ScheduledTask -TaskName $Name -TaskPath $TaskPath }
    'Remove' {
        $Task = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath
        if ($Task.State -in 'Running','Queued') { throw 'Wait for the active update to finish before removing the task.' }
        Unregister-ScheduledTask -TaskName $Name -TaskPath $TaskPath -Confirm:$false
    }
}
