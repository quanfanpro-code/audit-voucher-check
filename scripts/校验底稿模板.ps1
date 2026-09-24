# 仅在内存中使用虚构数据验证空白模板，不修改项目底稿或源模板。
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$template = Join-Path (Split-Path $PSScriptRoot -Parent) 'assets\抽凭检查底稿_空白模板.xlsx'
$hash = (Get-FileHash -LiteralPath $template).Hash
$excel = $null
$book = $null
$count = 0
function Check($condition, $message) {
    if (-not $condition) { throw $message }
    $script:count++
    Write-Output ('通过：' + $message)
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($template, 0, $true)
    $d = $book.Worksheets.Item('单张凭证底稿')
    $m = $book.Worksheets.Item('逐张抽凭记录')
    $f = $book.Worksheets.Item('五项标准复核')
    $excel.CalculateFull()
    Check ($book.Worksheets.Count -eq 5) 'Excel实际打开模板'
    Check (($m.Range('A5').Value2 -eq '凭证附件') -and ($m.Range('Q5').Value2 -eq '主体／账套') -and ($d.Range('A15').Value2 -eq '凭证附件')) '附件位于第一列且主体字段保留'
    Check (($m.Range('B6').Value2 -eq '') -and ($m.Range('E6').Value2 -eq '')) '空白原号不产生假凭证或零金额'
    Check ($d.Range('D5').NumberFormat -eq '@') '原始凭证号按文本保存'
    Check (($d.Range('K39').Validation.ShowError -eq $true) -and ($d.Range('K39').Validation.AlertStyle -eq 1)) '结果下拉启用停止型提示'
    # 程序写单元格可以绕过Excel输入提示，因此另校验Validation.Value。
    $d.Range('K39').Value2 = '发现异常'
    Check (-not $d.Range('K39').Validation.Value) '旧结果标签不能通过下拉有效性检查'
    $d.Range('K39').Value2 = '未执行'
    $d.Copy([Type]::Missing, $d)
    $second = $book.Worksheets.Item($d.Index + 1)
    $second.Name = '2099年01月记0002号'
    $d.Name = '2099年01月记0001号'
    Check ($m.Range('L6').Formula.Contains("#'单张凭证底稿'!A1")) '确认页签改名不会自动修复文字链接'
    foreach ($cell in @($m.Range('L6'), $f.Range('J6'))) {
        $cell.Formula = ([string]$cell.Formula).Replace('单张凭证底稿', $d.Name)
    }
    $d.Range('D4').Value2 = '虚构测试账套'
    $d.Range('D5').Value2 = $d.Name
    $d.Range('D9').Value2 = 2500.0
    $d.Range('D15').Value2 = '虚构测试：采购合同、付款审批单、银行回单。'
    $d.Range('K39').Value2 = '不正确'
    $d.Range('L39').Value2 = '虚构验证：数量多记5件、金额多记500；另有服务税额100的资格未明。'
    $d.Range('D55').Value2 = '仅服务税额100的抵扣资格未明；数量差错500已查明。'
    $excel.CalculateFull()
    Check (($m.Range('A6').Value2 -eq $d.Range('D15').Value2) -and ($m.Range('Q6').Value2 -eq '虚构测试账套')) '附件摘要和主体分别进入总表'
    Check ($m.Range('I6').Value2 -eq '未执行') '五项输入不会自动伪造整张专业结论'
    $d.Range('D52').Value2 = '发现明确问题，另有事项未能判断'
    $excel.CalculateFull()
    Check (($f.Range('E6').Value2 -eq '不正确') -and ($f.Range('I6').Value2 -like '*100*') -and ($m.Range('I6').Value2 -eq $d.Range('D52').Value2)) '已证实错误和具体限制同时进入汇总'
    $d.Range('D9').Value2 = 0.0
    $excel.CalculateFull()
    Check ($m.Range('E6').Value2 -ceq 0.0) '零金额和空白分开'
    $second.Range('D5').Value2 = $second.Name
    $second.Range('D15').Value2 = '虚构测试：费用报销单、出租车电子发票。'
    $second.Range('D52').Value2 = '所执行检查未发现具体问题'
    $m.Range('A6:Q6').Copy($m.Range('A7:Q7')) | Out-Null
    $f.Range('A6:J6').Copy($f.Range('A7:J7')) | Out-Null
    foreach ($cell in $m.Range('A7:Q7').Cells) {
        if ($cell.HasFormula) { $cell.Formula = ([string]$cell.Formula).Replace($d.Name, $second.Name) }
    }
    foreach ($cell in $f.Range('A7:J7').Cells) {
        if ($cell.HasFormula) { $cell.Formula = ([string]$cell.Formula).Replace($d.Name, $second.Name) }
    }
    $excel.CalculateFull()
    Check (($m.Range('B7').Value2 -eq $second.Name) -and ($m.Range('I7').Value2 -eq '所执行检查未发现具体问题') -and ($m.Range('K7').Value2 -eq '')) '第二张引用独立且不继承前张限制'
    Check (($m.Range('A7').Value2 -eq $second.Range('D15').Value2) -and ($m.Range('A7').Value2 -ne $m.Range('A6').Value2)) '两张附件摘要分别引用且不串用'
    foreach ($row in 6,7) {
        $target = [regex]::Match([string]$m.Range("L$row").Formula, "#'([^']+)'!A1").Groups[1].Value
        Check ($target -eq $m.Range("B$row").Value2) ('第' + $row + '行链接与完整凭证号对应')
        $excel.Goto($book.Worksheets.Item($target).Range('A1'))
        Check ($book.ActiveSheet.Name -eq $target) ('公式目标页签存在且可定位到' + $target)
        $back = $book.Worksheets.Item($target).Range('A1')
        Check (($back.Value2 -eq '返回逐张抽凭记录') -and ($back.Formula -eq '=HYPERLINK("#''逐张抽凭记录''!A5","返回逐张抽凭记录")')) '复制页保留有效返回总表链接'
        $excel.Goto($m.Range('A5'))
        Check ($book.ActiveSheet.Name -eq $m.Name) '总表页签存在且可定位'
    }
    $second.Range('25:26').EntireRow.Insert() | Out-Null
    $excel.CalculateFull()
    Check ($m.Range('I7').Value2 -eq '所执行检查未发现具体问题') '扩行后当前结论引用保持正确'
    # 同一原号可属于两个账套；页签用不同名称，表内保留原号和账套。
    $oldName = $second.Name
    $second.Name = '乙账套_2099年01月记0001号'
    $second.Range('D4').Value2 = '虚构测试乙账套'
    $second.Range('D5').Value2 = $d.Range('D5').Value2
    foreach ($cell in @($m.Range('L7'), $f.Range('J7'))) {
        $cell.Formula = ([string]$cell.Formula).Replace($oldName, $second.Name)
    }
    $excel.CalculateFull()
    Check (($m.Range('B6').Value2 -eq $m.Range('B7').Value2) -and ($m.Range('Q6').Value2 -ne $m.Range('Q7').Value2)) '同原号跨账套仍保留两条独立记录'
    $targetA = [regex]::Match([string]$m.Range('L6').Formula, "#'([^']+)'!A1").Groups[1].Value
    $targetB = [regex]::Match([string]$m.Range('L7').Formula, "#'([^']+)'!A1").Groups[1].Value
    Check (($targetA -ne $targetB) -and ($targetA -eq $d.Name) -and ($targetB -eq $second.Name)) '同原号跨账套的明细目标不串页'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
Check ((Get-FileHash -LiteralPath $template).Hash -eq $hash) '源模板未被测试数据改动'
Write-Output ('共通过 ' + $count + ' 项文件功能检查；这些检查不代替实际业务判断。')
