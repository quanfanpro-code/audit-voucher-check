# 128K 串行上下文窗口调度外壳（抽凭检查）：只管理占用、哈希与短状态，不是业务真相。
# 实际完成范围与下一张凭证仍以指定清单和已保存工作簿为准；本脚本不打开正式工作簿。
param(
    [Parameter(Mandatory = $true)][string]$Action,
    [string]$Root,
    [string]$Mode,
    [string]$ModelBinding,
    [string]$BackupRoot,
    [string]$TaskId,
    [string]$TaskPacket,
    [string]$WindowId,
    [string]$ResultWorkbook,
    [string]$Reason,
    [switch]$SavedAndReadBack
)
$ErrorActionPreference = 'Stop'
$ControllerDispatchLimit = 20
$stateDir = Join-Path $Root '.上下文调度'
$stateFile = Join-Path $stateDir '上下文调度状态.json'
$logFile = Join-Path $stateDir '上下文调度记录.jsonl'

function New-调度异常([string]$message) {
    return (New-Object -TypeName System.InvalidOperationException -ArgumentList $message)
}

function Get-文件哈希([string]$path) {
    return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
}

function Backup-调度文件([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return }
    if ([string]::IsNullOrWhiteSpace($BackupRoot)) { throw (New-调度异常 '缺少 -BackupRoot，覆盖调度文件前必须提供备份目录') }
    $relative = $path.Substring($Root.TrimEnd('\').Length).TrimStart('\')
    $target = Join-Path $BackupRoot $relative
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $path -Destination $target -Force
    if ((Get-FileHash -LiteralPath $target).Hash -ne (Get-FileHash -LiteralPath $path).Hash) {
        throw (New-调度异常 ('备份字节校验失败：' + (Split-Path $path -Leaf)))
    }
}

function Save-调度状态($state) {
    Backup-调度文件 $stateFile
    $tmp = "$stateFile.$PID.tmp"
    $json = $state | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($tmp, $json, [Text.UTF8Encoding]::new($true))
    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        try {
            Move-Item -LiteralPath $tmp -Destination $stateFile -Force
            return
        } catch {
            if ($attempt -eq 4) { throw }
            Start-Sleep -Milliseconds (50 * ($attempt + 1))
        }
    }
}

function Add-调度日志([string]$event, $detail) {
    Backup-调度文件 $logFile
    $entryObject = [ordered]@{ event = $event; time = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); detail = $detail }
    $line = $entryObject | ConvertTo-Json -Depth 6 -Compress
    if (Test-Path -LiteralPath $logFile) {
        [IO.File]::AppendAllText($logFile, $line + "`n", [Text.UTF8Encoding]::new($false))
    } else {
        [IO.File]::WriteAllText($logFile, $line + "`n", [Text.UTF8Encoding]::new($true))
    }
}

function Read-调度状态 {
    if (-not (Test-Path -LiteralPath $stateFile)) { throw (New-调度异常 '尚未初始化 128K 调度状态') }
    $state = [IO.File]::ReadAllText($stateFile) | ConvertFrom-Json
    if ($state.context_mode -ne '128K') { throw (New-调度异常 '当前运行不是 128K 模式，禁止调度') }
    return $state
}

switch ($Action) {
    '初始化' {
        if ($Mode -eq '1M') { return }
        if ($Mode -ne '128K') { throw (New-调度异常 ('不支持的上下文模式：' + $Mode)) }
        if (-not (Test-Path -LiteralPath $stateFile)) {
            New-Item -ItemType Directory -Path $stateDir -Force | Out-Null
            $state = [ordered]@{
                context_mode = '128K'
                model_binding = $ModelBinding
                active_dispatch = $null
                controller_dispatch_count = 0
                workbook_anchor = $null
                created_at = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            }
            Save-调度状态 $state
            Add-调度日志 '初始化' @{ model_binding = $ModelBinding }
        }
    }
    '开始' {
        $state = Read-调度状态
        if ($null -ne $state.active_dispatch) { throw (New-调度异常 '存在活动任务，禁止开始第二张凭证') }
        if ([string]::IsNullOrWhiteSpace($TaskId)) { throw (New-调度异常 '缺少完整凭证键 -TaskId') }
        if (-not (Test-Path -LiteralPath $TaskPacket)) { throw (New-调度异常 '原任务文件不存在') }
        if ((-not [string]::IsNullOrEmpty($ModelBinding)) -and ($state.model_binding -ne $ModelBinding)) {
            throw (New-调度异常 '模型绑定不一致，禁止更换模型')
        }
        $packet = [IO.File]::ReadAllText($TaskPacket) | ConvertFrom-Json
        if (($packet.PSObject.Properties.Name -contains 'voucher_key') -and ($packet.voucher_key -ne $TaskId)) {
            throw (New-调度异常 ('任务编号与任务包凭证键不一致：' + $TaskId))
        }
        $state.active_dispatch = [ordered]@{
            task_id = $TaskId
            window_id = $WindowId
            packet_path = $TaskPacket
            packet_sha256 = Get-文件哈希 $TaskPacket
            started_at = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
        Save-调度状态 $state
        Add-调度日志 'start' $state.active_dispatch
    }
    '完成' {
        $state = Read-调度状态
        $active = $state.active_dispatch
        if ($null -eq $active) { throw (New-调度异常 '没有活动任务可登记完成') }
        if ($active.task_id -ne $TaskId) { throw (New-调度异常 ('任务编号与活动任务不一致：' + $TaskId + ' ≠ ' + $active.task_id)) }
        if (-not $SavedAndReadBack) { throw (New-调度异常 '未明确声明保存并回读（-SavedAndReadBack），不得完成') }
        if ([string]::IsNullOrWhiteSpace($ResultWorkbook)) { throw (New-调度异常 '缺少 -ResultWorkbook') }
        if (-not (Test-Path -LiteralPath $ResultWorkbook)) { throw (New-调度异常 '结果工作簿不存在') }
        $detail = [ordered]@{
            task_id = $TaskId
            window_id = $active.window_id
            workbook_path = $ResultWorkbook
            workbook_sha256 = Get-文件哈希 $ResultWorkbook
        }
        $state.workbook_anchor = [ordered]@{ path = $ResultWorkbook; sha256 = $detail.workbook_sha256 }
        $state.active_dispatch = $null
        $state.controller_dispatch_count = [int]$state.controller_dispatch_count + 1
        Save-调度状态 $state
        Add-调度日志 '完成' $detail
    }
    '状态' {
        $state = Read-调度状态
        $count = [int]$state.controller_dispatch_count
        $summary = [ordered]@{
            active_dispatch = $state.active_dispatch
            completed_in_controller = $count
            rotation_required = ($count -ge $ControllerDispatchLimit)
            workbook_anchor = $state.workbook_anchor
        }
        Write-Output ($summary | ConvertTo-Json -Depth 6)
    }
    '交接' {
        $state = Read-调度状态
        $count = [int]$state.controller_dispatch_count
        $info = [ordered]@{
            context_mode = $state.context_mode
            model_binding = $state.model_binding
            active_dispatch = $state.active_dispatch
            workbook_anchor = $state.workbook_anchor
            rotation_required = ($count -ge $ControllerDispatchLimit)
        }
        if ($count -gt 0) {
            $state.controller_dispatch_count = 0
            Save-调度状态 $state
            Add-调度日志 '主控轮换' @{ handed_off = $true }
        }
        Write-Output ($info | ConvertTo-Json -Depth 6)
    }
    '恢复' {
        $state = Read-调度状态
        $active = $state.active_dispatch
        if ($null -eq $active) { throw (New-调度异常 '没有活动任务占用，无需恢复') }
        if ($active.task_id -ne $TaskId) { throw (New-调度异常 ('任务编号与活动任务不一致：' + $TaskId + ' ≠ ' + $active.task_id)) }
        $state.active_dispatch = $null
        Save-调度状态 $state
        Add-调度日志 '恢复' @{ task_id = $TaskId; window_id = $active.window_id; reason = $Reason }
    }
    default { throw (New-调度异常 ('不支持的动作：' + $Action)) }
}
