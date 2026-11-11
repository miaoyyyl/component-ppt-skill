param(
    [ValidateRange(1024,65535)][int]$Port = 8871,
    [switch]$OpenBrowser,
    [switch]$Stop
)
$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$baseUrl = "http://127.0.0.1:$Port"
$runtimePath = Join-Path $projectRoot '.runtime'
New-Item -ItemType Directory -Path $runtimePath -Force | Out-Null
function Get-TaskService {
    try { return Invoke-RestMethod -Uri "$baseUrl/api/bootstrap" -TimeoutSec 2 } catch { return $null }
}
$serviceInfo = Get-TaskService
if ($Stop) {
    if ($null -eq $serviceInfo -or $serviceInfo.applicationId -ne 'ppt-component-task-server' -or $serviceInfo.workspace -ne $projectRoot) { throw '没有找到本项目的任务入口。' }
    Invoke-RestMethod -Method Post -Uri "$baseUrl/api/shutdown" -Headers @{ Origin = $baseUrl; 'X-Task-Token' = $serviceInfo.token } -ContentType 'application/json' -Body '{}' | Out-Null
    Write-Host '本地任务入口已停止。'
    exit
}
if ($null -eq $serviceInfo) {
    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($null -eq $pythonCommand) { throw '没有找到Python，请配置Python后重新启动。' }
    $serverScript = Join-Path $projectRoot 'scripts/task_server.py'
    Start-Process -FilePath $pythonCommand.Source -ArgumentList @('"' + $serverScript + '"', '--port', "$Port") -WorkingDirectory $projectRoot -WindowStyle Hidden -RedirectStandardOutput (Join-Path $runtimePath 'server-stdout.log') -RedirectStandardError (Join-Path $runtimePath 'server-stderr.log') | Out-Null
    for ($attempt=0; $attempt -lt 80; $attempt++) {
        Start-Sleep -Milliseconds 250
        $serviceInfo = Get-TaskService
        if ($null -ne $serviceInfo) { break }
    }
}
if ($null -eq $serviceInfo -or $serviceInfo.applicationId -ne 'ppt-component-task-server' -or $serviceInfo.workspace -ne $projectRoot) { throw '任务入口未能启动，请查看 .runtime/server-stderr.log；端口也可能被占用。' }
$pageUrl = "$baseUrl/library/index.html"
Write-Host "制作任务入口：$pageUrl"
Write-Host $serviceInfo.runner.message
if ($OpenBrowser) { Start-Process -FilePath $pageUrl -WindowStyle Normal | Out-Null }
