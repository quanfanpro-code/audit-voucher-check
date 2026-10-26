# 调度外壳功能自检：只使用虚构资料和临时目录，不读取真实凭证，不打开正式工作簿。
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$dispatcher = Join-Path (Split-Path $PSScriptRoot -Parent) 'scripts\串行上下文调度.ps1'
$root = Join-Path $env:TEMP ('抽凭调度测试_' + [guid]::NewGuid().ToString('N'))
$backup = Join-Path $root 'backup'
$task = Join-Path $root '单张凭证任务.json'
New-Item -ItemType Directory -Path $root -Force | Out-Null
$packetObject = [ordered]@{
    voucher_key = '测试账套|2026|01|记|0001'
    workbook_path = (Join-Path $root '测试账套_抽凭检查底稿.xlsx')
    approved_scope = @('测试账套|2026|01|记|0001')
    source_indexes = @('虚构账簿索引.json', '虚构附件索引.json')
    source_versions = @{ algorithm = 'SHA256'; 账簿索引 = ('0' * 64); 附件索引 = ('0' * 64) }
    required_references = @('SKILL.md', 'references/判断边界与实例.md')
    check_type = '首次检查'
    output_location = (Join-Path $root '测试账套_抽凭检查底稿.xlsx')
}
[IO.File]::WriteAllText($task, ($packetObject | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($true))

& $dispatcher -Action 初始化 -Root $root -Mode 1M -ModelBinding current-model -BackupRoot $backup
if (Test-Path (Join-Path $root '.上下文调度\上下文调度状态.json')) { throw '1M不应创建调度状态' }
& $dispatcher -Action 初始化 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup
& $dispatcher -Action 开始 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup -TaskId '测试账套|2026|01|记|0001' -TaskPacket $task -WindowId window-A
$rejected = $false
try {
    & $dispatcher -Action 开始 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup -TaskId '测试账套|2026|01|记|0002' -TaskPacket $task -WindowId window-B
} catch [InvalidOperationException] { $rejected = $true }
if (-not $rejected) { throw '活动凭证未解除时第二凭证不得开始' }
Write-Output '通过：1M旁路且128K串行占用'

$workbook = Join-Path $root '测试账套_抽凭检查底稿.xlsx'
[IO.File]::WriteAllBytes($workbook, [byte[]](1, 2, 3))
$completionRejected = $false
try {
    & $dispatcher -Action 完成 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup -TaskId '测试账套|2026|01|记|0001' -ResultWorkbook $workbook
} catch [InvalidOperationException] { $completionRejected = $true }
if (-not $completionRejected) { throw '未明确保存并回读不得完成' }
& $dispatcher -Action 完成 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup -TaskId '测试账套|2026|01|记|0001' -ResultWorkbook $workbook -SavedAndReadBack
$state = & $dispatcher -Action 状态 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup | ConvertFrom-Json
if ($null -ne $state.active_dispatch) { throw '完成后必须解除占用' }
Write-Output '通过：保存并回读是完成门禁'

for ($index = 2; $index -le 20; $index++) {
    $key = '测试账套|2026|01|记|{0:D4}' -f $index
    $roundTask = Join-Path $root ('单张凭证任务_' + $index + '.json')
    $packetObject.voucher_key = $key
    [IO.File]::WriteAllText($roundTask, ($packetObject | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($true))
    $roundBook = Join-Path $root ('结果_' + $index + '.xlsx')
    Copy-Item -LiteralPath $workbook -Destination $roundBook
    & $dispatcher -Action 开始 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup -TaskId $key -TaskPacket $roundTask -WindowId ('window-' + $index)
    & $dispatcher -Action 完成 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup -TaskId $key -ResultWorkbook $roundBook -SavedAndReadBack
}
$state = & $dispatcher -Action 状态 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup | ConvertFrom-Json
if (-not $state.rotation_required) { throw '完成20次调度后必须轮换主控窗口' }
Write-Output '通过：主控窗口轮换'

# —— Task 3：双入口与单张任务包契约 ——
$skillRoot = Split-Path $PSScriptRoot -Parent
$skill = [IO.File]::ReadAllText((Join-Path $skillRoot 'SKILL.md'), [Text.Encoding]::UTF8)
$guide = [IO.File]::ReadAllText((Join-Path $skillRoot 'references\128K串行上下文.md'), [Text.Encoding]::UTF8)
if (-not $skill.Contains('本次使用 1M 还是 128K 上下文模型')) { throw '缺少上下文入口问题' }
foreach ($phrase in @('前台', '串行', '同一模型', 'SavedAndReadBack', '完整凭证键', '不得预设结论')) {
    if (-not $guide.Contains($phrase)) { throw ('128K说明缺少：' + $phrase) }
}
$packet = Get-Content -LiteralPath $task -Raw -Encoding UTF8 | ConvertFrom-Json
$requiredPacketFields = @('voucher_key', 'workbook_path', 'approved_scope', 'source_indexes', 'source_versions', 'required_references', 'check_type', 'output_location')
foreach ($field in $requiredPacketFields) {
    if ($packet.PSObject.Properties.Name -notcontains $field) { throw ('任务包缺少字段：' + $field) }
}
if ($packet.PSObject.Properties.Name -contains 'conclusion') { throw '调度任务包不得预设专业结论' }
$prompts = Get-Content -LiteralPath (Join-Path $skillRoot 'test-prompts.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$originalIds = @('付款凭据与后附范围', '设备验收与金额', '按旧底稿定点重判')
for ($i = 0; $i -lt 3; $i++) {
    if ($prompts[$i].id -ne $originalIds[$i]) { throw ('既有专业情景顺序被改动：' + $prompts[$i].id) }
    foreach ($fieldName in @('id','prompt','expected')) {
        if ([string]::IsNullOrWhiteSpace([string]$prompts[$i].$fieldName)) { throw ('既有情景字段缺失：' + $originalIds[$i] + '.' + $fieldName) }
    }
}
if ($prompts.Count -lt 5) { throw 'test-prompts.json 未追加128K上下文情景' }
Write-Output '通过：双入口与单张任务包契约'

# —— Task 4：断点恢复与短交接 ——
$recoveryKey = '测试账套|2026|01|记|0099'
$recoveryTask = Join-Path $root '单张凭证任务_恢复.json'
$packetObject.voucher_key = $recoveryKey
[IO.File]::WriteAllText($recoveryTask, ($packetObject | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($true))
& $dispatcher -Action 开始 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup -TaskId $recoveryKey -TaskPacket $recoveryTask -WindowId window-recovery
$workbookHashBefore = (Get-FileHash -LiteralPath $workbook -Algorithm SHA256).Hash
& $dispatcher -Action 恢复 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup -TaskId $recoveryKey -Reason '模拟执行窗口中断'
$state = & $dispatcher -Action 状态 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup | ConvertFrom-Json
if ($null -ne $state.active_dispatch) { throw '恢复后仍占用凭证' }
if ((Get-FileHash -LiteralPath $workbook -Algorithm SHA256).Hash -ne $workbookHashBefore) { throw '恢复改写了工作簿' }
$handoffText = (& $dispatcher -Action 交接 -Root $root -Mode 128K -ModelBinding current-model -BackupRoot $backup | Out-String)
foreach ($forbidden in @('conclusion', 'evidence_text', 'ocr_text')) {
    if ($handoffText.Contains($forbidden)) { throw ('交接不得包含专业正文：' + $forbidden) }
}
Write-Output '通过：断点恢复与短交接'
