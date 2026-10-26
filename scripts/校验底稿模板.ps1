# 仅在内存中使用虚构数据验证空白模板（2026-10-22 三页签版式），不修改项目底稿或源模板。
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
    Check ($book.Worksheets.Count -eq 3) 'Excel实际打开模板且为三页签'
    $s = $book.Worksheets.Item('说明')
    $m = $book.Worksheets.Item('抽凭结果总表')
    $d = $book.Worksheets.Item('单张凭证底稿')
    Check (($s.Range('A1').Value2 -eq '项目名称') -and ($s.Range('A6').Value2 -eq '结论口径') -and ($s.Range('A7').Value2 -eq '资料形态与检查限制')) '说明页承载全簿共同口径且不重复'
    $expectedHeaders = @('凭证号','源文件','记账日期','业务内容','受检金额(元)','原始凭证','账证相符','账务处理','期间归属','回款户名','一句话结论','问题情况','明细页','后附附件概要')
    $actualHeaders = @()
    foreach ($cell in $m.Range('A1:N1').Cells) { $actualHeaders += [string]$cell.Value2 }
    $headerDiff = Compare-Object -ReferenceObject $expectedHeaders -DifferenceObject $actualHeaders
    Check (($null -eq $headerDiff) -or ($headerDiff.Count -eq 0)) '总表表头列名与顺序符合第六节'
    Check ($m.Range('A1').Font.Bold -eq $true) '总表表头加粗'
    $m.Activate()
    Check ($excel.ActiveWindow.FreezePanes -and $excel.ActiveWindow.SplitRow -eq 1) '总表只冻结表头一行'
    $d.Activate()
    Check ($excel.ActiveWindow.FreezePanes -and $excel.ActiveWindow.SplitRow -eq 1) '单张底稿冻结首行'
    $s.Activate()
    Check (-not $excel.ActiveWindow.FreezePanes) '说明页无表头不冻结'
    Check (($d.Range('H1').Formula -like '*HYPERLINK*#抽凭结果总表*') -and ($d.Range('H1').Value2 -eq '← 返回总表')) '单张底稿标题行右端保留返回总表链接'
    Check ([string]$m.Range('M2').Formula).Contains("#'单张凭证底稿'!A1") '总表明细页链接指向单张底稿'
    foreach ($pair in @(@('F2','齐全'), @('G2','相符'), @('H2','正确'), @('I2','恰当'), @('J2','一致'), @('L2','明确问题'))) {
        $v = $m.Range($pair[0]).Validation
        Check (($v.Type -eq 3) -and ([string]$v.Formula1).Contains($pair[1])) ('总表' + $pair[0] + '下拉选项匹配新口径（' + $pair[1] + '）')
    }
    foreach ($pair in @(@('H34','齐全'), @('H35','相符'), @('H36','正确'), @('H37','恰当'), @('H38','一致'))) {
        $v = $d.Range($pair[0]).Validation
        Check (($v.Type -eq 3) -and ([string]$v.Formula1).Contains($pair[1])) ('底稿' + $pair[0] + '下拉选项与总表一致（' + $pair[1] + '）')
    }
    $v = $d.Range('B42').Validation
    Check (($v.Type -eq 3) -and ([string]$v.Formula1).Contains('未执行') -and ([string]$v.Formula1).Contains('发现明确问题，另有事项未能判断')) '当前结论下拉为六种新口径'
    $d.Range('B42').Value2 = '旧结果标签'
    Check (-not $d.Range('B42').Validation.Value) '旧结果标签不在当前结论下拉选项内'
    $d.Range('B42').Value2 = '未执行'
    $excel.CalculateFull()
    Check ($m.Range('A2').Value2 -eq '（凭证号）') '总表凭证号公式连通并显示底稿占位文本'
    Check ($m.Range('K2').Value2 -eq '未执行') '当前结论初始状态为未执行并连通总表'
    Check ($d.Range('B8').NumberFormat -eq '#,##0.00') '受检金额为千分位两位小数'
    $d.Range('B4').Value2 = '2099年1月第0001号'
    $d.Range('B5').Value2 = '2099-01-05'
    $d.Range('B6').Value2 = '虚构测试：支付办公设备款。'
    $d.Range('B8').Value2 = 2500.0
    $d.Range('H34').Value2 = '齐全'
    $d.Range('H36').Value2 = '不正确'
    $excel.CalculateFull()
    Check ([string]$m.Range('K2').Value2 -eq '未执行') '五项输入不会自动伪造整张专业结论'
    $d.Range('B42').Value2 = '发现明确问题，另有事项未能判断'
    $excel.CalculateFull()
    Check ($m.Range('A2').Value2 -eq '2099年1月第0001号') '凭证号进入总表'
    Check ($m.Range('E2').Value2 -eq 2500.0) '受检金额进入总表'
    Check (($m.Range('F2').Value2 -eq '齐全') -and ($m.Range('H2').Value2 -eq '不正确')) '五项结果进入总表对应列'
    Check ($m.Range('K2').Value2 -eq '发现明确问题，另有事项未能判断') '当前结论进入总表'
    $d.Copy([Type]::Missing, $d)
    $second = $book.Worksheets.Item($d.Index + 1)
    $second.Name = '2099年1月第0002号'
    $d.Name = '2099年1月第0001号'
    Check ([string]$m.Range('M2').Formula).Contains("#'单张凭证底稿'!A1") '页签改名不会自动修复文字链接'
    $m.Range('M2').Formula = ([string]$m.Range('M2').Formula).Replace('单张凭证底稿', $d.Name)
    $m.Range('A3').Formula = "=IF('$($second.Name)'!`$B4=`"`",`"`",'$($second.Name)'!`$B4)"
    $m.Range('C3').Formula = "=IF('$($second.Name)'!`$B5=`"`",`"`",'$($second.Name)'!`$B5)"
    $m.Range('D3').Formula = "=IF('$($second.Name)'!`$B6=`"`",`"`",'$($second.Name)'!`$B6)"
    $m.Range('E3').Formula = "=IF('$($second.Name)'!`$B8=`"`",`"`",'$($second.Name)'!`$B8)"
    foreach ($pair in @(@('F3','H34'), @('G3','H35'), @('H3','H36'), @('I3','H37'), @('J3','H38'), @('K3','B42'), @('N3','B14'))) {
        $m.Range($pair[0]).Formula = "=IF('$($second.Name)'!`$$($pair[1])=`"`",`"`",'$($second.Name)'!`$$($pair[1]))"
    }
    $m.Range('M3').Formula = "=IF(`$A`$3=`"`",`"`",HYPERLINK(`"#'$($second.Name)'!A1`"`,`"打开`"))"
    $second.Range('B4').Value2 = '2099年1月第0002号'
    $second.Range('H34').Value2 = '齐全'
    $second.Range('B42').Value2 = '所执行检查未发现具体问题'
    $excel.CalculateFull()
    Check ($m.Range('A2').Value2 -eq '2099年1月第0001号') '第一张引用随改名保持正确'
    Check (($m.Range('A3').Value2 -eq '2099年1月第0002号') -and ($m.Range('F3').Value2 -eq '齐全') -and ($m.Range('K3').Value2 -eq '所执行检查未发现具体问题')) '第二张引用独立且结果正确'
    Check ($m.Range('K3').Value2 -ne $m.Range('K2').Value2) '两张结论互不串用'
    foreach ($row in @(2, 3)) {
        $linkFormula = [string]$m.Range("M$row").Formula
        $target = [regex]::Match($linkFormula, "#'([^']+)'!A1").Groups[1].Value
        Check ($null -ne $book.Worksheets.Item($target)) ('第' + $row + '行明细链接目标页签存在')
        $excel.Goto($book.Worksheets.Item($target).Range('A1'))
        Check ($book.ActiveSheet.Name -eq $target) ('第' + $row + '行链接可定位到' + $target)
    }
    $excel.Goto($m.Range('A1'))
    Check ($book.ActiveSheet.Name -eq $m.Name) '总表页签存在且可定位'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
Check ((Get-FileHash -LiteralPath $template).Hash -eq $hash) '源模板未被测试数据改动'
Write-Output ('共通过 ' + $count + ' 项文件功能检查；这些检查不代替实际业务判断。')
