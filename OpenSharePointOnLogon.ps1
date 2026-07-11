$taskName = "OpenSharePointOnLogon"
$xmlPath = "C:\ProgramData\CompanyScripts\OpenSharePoint.xml"
$scriptDir = "C:\ProgramData\CompanyScripts"

# フォルダ作成
if (!(Test-Path $scriptDir)) {
    New-Item -Path $scriptDir -ItemType Directory -Force | Out-Null
}

# XML タスクを書き出し
$xmlContent = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Open SharePoint on Friday morning logon</Description>
  </RegistrationInfo>

  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>

      <!-- 朝6時〜10時の間だけ有効 -->
      <StartBoundary>2025-01-01T06:00:00</StartBoundary>
      <EndBoundary>2025-01-01T10:00:00</EndBoundary>

      <!-- 毎週金曜のみ -->
      <ScheduleByWeek>
        <DaysOfWeek>
          <Friday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>

      <!-- 15秒遅延 -->
      <Delay>PT15S</Delay>
    </LogonTrigger>
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
      <Command>"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"</Command>

      <!-- アプリモード + ウィンドウサイズ + 画面中央配置 -->
      <Arguments>--app="https://exmaple.com" --window-size=800,600 --window-position=112,84</Arguments>
    </Exec>
  </Actions>
</Task>
'@
$xmlContent | Out-File -FilePath $xmlPath -Encoding UTF8 -Force

# 既存タスク削除
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# XML からタスク登録
Register-ScheduledTask -TaskName $taskName -Xml (Get-Content $xmlPath | Out-String)
