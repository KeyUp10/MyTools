$taskName = "OpenSharePointOnLogon"
$xmlPath = "C:\ProgramData\CompanyScripts\OpenSharePoint.xml"
$ps1Path = "C:\ProgramData\CompanyScripts\OpenSharePoint.ps1"
$scriptDir = "C:\ProgramData\CompanyScripts"

# フォルダ作成
if (!(Test-Path $scriptDir)) {
    New-Item -Path $scriptDir -ItemType Directory -Force | Out-Null
}

# PowerShell スクリプト書き出し
$ps1Content = @"
$hour = (Get-Date).Hour

if ($hour -ge 6 -and $hour -le 10) {
    Start-Process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
        "--app=https://exmaple.com --window-size=800,600 --window-position=112,84"
}
"@
$ps1Content | Out-File -FilePath $ps1Path -Encoding UTF8 -Force

# XML タスク書き出し（@" ～ "@ で安全）
$xmlContent = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Open SharePoint on morning logon or unlock</Description>
  </RegistrationInfo>

  <Triggers>
    <!-- ログオン時 -->
    <LogonTrigger>
      <Enabled>true</Enabled>
      <Delay>PT15S</Delay>
    </LogonTrigger>

    <!-- 再接続（SessionUnlock） -->
    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <StateChange>SessionUnlock</StateChange>
      <Delay>PT15S</Delay>
    </SessionStateChangeTrigger>
  </Triggers>

  <Principals>
    <Principal id="Users">
      <GroupId>S-1-5-32-545</GroupId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>

  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>

  <Actions Context="Users">
    <Exec>
      <Command>"powershell.exe"</Command>
      <Arguments>-ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\ProgramData\CompanyScripts\OpenSharePoint.ps1"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

$xmlContent | Out-File -FilePath $xmlPath -Encoding UTF8 -Force

# 既存タスク削除
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# XML からタスク登録
Register-ScheduledTask -TaskName $taskName -Xml (Get-Content $xmlPath | Out-String)
