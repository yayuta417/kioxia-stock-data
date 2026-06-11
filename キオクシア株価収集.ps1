# ============================================================
# キオクシアホールディングス (285A.T) 株価データ収集スクリプト
# ------------------------------------------------------------
# データ源 : Yahoo Finance チャートAPI (認証不要)
# 取得内容 : 1分足(過去28日を7日×4回に分割取得) / 5分足(過去60日)
# 保存先   : このスクリプトと同じフォルダ
#   1分足\285A_1分足_YYYY-MM-DD.csv   (時刻,始値,高値,安値,終値,出来高)
#   5分足\285A_5分足_YYYY-MM-DD.csv   (同上)
#   日次サマリー.csv (日付,始値,前場引値,後場始値,終値,高値,安値,出来高)
#   収集ログ.txt
# 仕様     : 取得できた日付のCSVは毎回上書き(自己修復)。
#            前場引値 = 11:30以前の最後のバーの終値
#            後場始値 = 12:30以降の最初のバーの始値
# 終了コード: 0 = 正常(休場日含む) / 1 = 1分足が1件も取得できない異常
# ============================================================

$ErrorActionPreference = "Stop"
# 保存先 = スクリプト自身のフォルダ(ローカル/GitHub Actions どちらでも動作)
$baseDir = $PSScriptRoot
if (-not $baseDir) { $baseDir = "C:\Users\user\.claude\claude_code_作業用\キオクシアのデータ" }
$symbol  = "285A.T"
$jst     = [TimeZoneInfo]::FindSystemTimeZoneById("Tokyo Standard Time")
$epoch   = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
$logPath = Join-Path $baseDir "収集ログ.txt"

function Write-Log {
    param($msg)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logPath -Value $line -Encoding UTF8
    Write-Output $line
}

function Get-Chart {
    param($interval, $p1, $p2, $range)
    if ($range) {
        $url = "https://query1.finance.yahoo.com/v8/finance/chart/{0}?interval={1}&range={2}" -f $symbol, $interval, $range
    } else {
        $url = "https://query1.finance.yahoo.com/v8/finance/chart/{0}?interval={1}&period1={2}&period2={3}" -f $symbol, $interval, $p1, $p2
    }
    try {
        $r = Invoke-RestMethod -Uri $url -Headers @{ "User-Agent" = "Mozilla/5.0" } -TimeoutSec 60
        return $r.chart.result[0]
    } catch {
        Write-Log ("API取得失敗 ({0} / {1}): {2}" -f $interval, $url, $_.Exception.Message)
        return $null
    }
}

function Format-Price {
    param($v)
    if ($null -eq $v) { return "" }
    $d = [math]::Round([double]$v, 2)
    if ($d -eq [math]::Floor($d)) { return [string][long]$d }
    return [string]$d
}

# APIレスポンスを日付(JST)ごとのバー一覧に変換する
function Get-BarsByDate {
    param($res)
    $byDate = @{}
    if ($null -eq $res -or $null -eq $res.timestamp) { return $byDate }
    $q = $res.indicators.quote[0]
    for ($i = 0; $i -lt $res.timestamp.Count; $i++) {
        $o = $q.open[$i]; $h = $q.high[$i]; $l = $q.low[$i]; $c = $q.close[$i]; $v = $q.volume[$i]
        if ($null -eq $o -and $null -eq $c) { continue }   # 昼休み等の空バーを除外
        $t = [TimeZoneInfo]::ConvertTimeFromUtc($epoch.AddSeconds($res.timestamp[$i]), $jst)
        $dateKey = $t.ToString("yyyy-MM-dd")
        if (-not $byDate.ContainsKey($dateKey)) { $byDate[$dateKey] = New-Object System.Collections.ArrayList }
        [void]$byDate[$dateKey].Add([PSCustomObject]@{ Time = $t; O = $o; H = $h; L = $l; C = $c; V = $v })
    }
    return $byDate
}

function Merge-BarsByDate {
    param($target, $chunk)
    foreach ($key in $chunk.Keys) {
        if (-not $target.ContainsKey($key)) { $target[$key] = New-Object System.Collections.ArrayList }
        foreach ($b in $chunk[$key]) { [void]$target[$key].Add($b) }
    }
}

function Save-Bars {
    param($byDate, $subdir, $label)
    $dir = Join-Path $baseDir $subdir
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $saved = New-Object System.Collections.ArrayList
    foreach ($dateKey in ($byDate.Keys | Sort-Object)) {
        $bars = @($byDate[$dateKey] | Sort-Object Time -Unique)
        if ($bars.Count -eq 0) { continue }
        $lines = New-Object System.Collections.ArrayList
        [void]$lines.Add("時刻,始値,高値,安値,終値,出来高")
        foreach ($b in $bars) {
            $vol = 0; if ($null -ne $b.V) { $vol = [long]$b.V }
            [void]$lines.Add(("{0},{1},{2},{3},{4},{5}" -f $b.Time.ToString("HH:mm"), (Format-Price $b.O), (Format-Price $b.H), (Format-Price $b.L), (Format-Price $b.C), $vol))
        }
        $file = Join-Path $dir ("285A_{0}_{1}.csv" -f $label, $dateKey)
        Set-Content -Path $file -Value $lines -Encoding UTF8
        [void]$saved.Add($dateKey)
    }
    return $saved.ToArray()
}

# 1分足から日次サマリー(始値・前場引値・後場始値・終値ほか)を算出する
function Get-DailySummary {
    param($dateKey, $bars)
    $valid = @($bars | Sort-Object Time -Unique)
    if ($valid.Count -eq 0) { return $null }
    $firstO = @($valid | Where-Object { $null -ne $_.O })
    $lastC  = @($valid | Where-Object { $null -ne $_.C })
    if ($firstO.Count -eq 0 -or $lastC.Count -eq 0) { return $null }
    $am = @($valid | Where-Object { $_.Time.TimeOfDay -le [TimeSpan]"11:30:00" -and $null -ne $_.C })
    $pm = @($valid | Where-Object { $_.Time.TimeOfDay -ge [TimeSpan]"12:30:00" -and $null -ne $_.O })
    $amClose = ""; if ($am.Count -gt 0) { $amClose = Format-Price $am[$am.Count - 1].C }
    $pmOpen  = ""; if ($pm.Count -gt 0) { $pmOpen = Format-Price $pm[0].O }
    $high = ($valid | Where-Object { $null -ne $_.H } | Measure-Object -Property H -Maximum).Maximum
    $low  = ($valid | Where-Object { $null -ne $_.L } | Measure-Object -Property L -Minimum).Minimum
    $volSum = ($valid | Where-Object { $null -ne $_.V } | Measure-Object -Property V -Sum).Sum
    $volStr = "0"; if ($null -ne $volSum) { $volStr = [string][long]$volSum }
    return [PSCustomObject]@{
        日付 = $dateKey
        始値 = (Format-Price $firstO[0].O)
        前場引値 = $amClose
        後場始値 = $pmOpen
        終値 = (Format-Price $lastC[$lastC.Count - 1].C)
        高値 = (Format-Price $high)
        安値 = (Format-Price $low)
        出来高 = $volStr
    }
}

# ---------------- メイン処理 ----------------
try {
    if (-not (Test-Path $baseDir)) { New-Item -ItemType Directory -Force -Path $baseDir | Out-Null }
    Write-Log "===== 収集開始 ====="

    # 1分足: 7日×4チャンクで過去28日分を取得(Yahooの1分足は約30日までしか遡れない)
    $byDate1m = @{}
    $nowUnix = [long]([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
    for ($k = 0; $k -lt 4; $k++) {
        $p2 = $nowUnix - ($k * 7 * 86400)
        $p1 = $p2 - (7 * 86400)
        $res = Get-Chart "1m" $p1 $p2 $null
        Merge-BarsByDate $byDate1m (Get-BarsByDate $res)
    }
    $saved1m = @(Save-Bars $byDate1m "1分足" "1分足")
    Write-Log ("1分足: {0}日分を保存 ({1})" -f $saved1m.Count, ($saved1m -join ", "))

    # 5分足: 60日分(失敗時は範囲を狭めて再試行)
    $res5 = Get-Chart "5m" $null $null "60d"
    if ($null -eq $res5) { $res5 = Get-Chart "5m" $null $null "1mo" }
    if ($null -eq $res5) { $res5 = Get-Chart "5m" $null $null "7d" }
    $byDate5m = Get-BarsByDate $res5
    $saved5m = @(Save-Bars $byDate5m "5分足" "5分足")
    Write-Log ("5分足: {0}日分を保存" -f $saved5m.Count)

    # 日次サマリーの更新(既存行とマージし日付順に書き戻す)
    $summaryPath = Join-Path $baseDir "日次サマリー.csv"
    $existing = @{}
    if (Test-Path $summaryPath) {
        foreach ($row in @(Import-Csv $summaryPath)) { $existing[$row.日付] = $row }
    }
    # 1分足が取得できない日(約30日より前)は5分足から日次サマリーを補完
    foreach ($dateKey in $byDate5m.Keys) {
        if ($byDate1m.ContainsKey($dateKey) -or $existing.ContainsKey($dateKey)) { continue }
        $s = Get-DailySummary $dateKey $byDate5m[$dateKey]
        if ($null -ne $s) { $existing[$dateKey] = $s }
    }
    foreach ($dateKey in $byDate1m.Keys) {
        $s = Get-DailySummary $dateKey $byDate1m[$dateKey]
        if ($null -ne $s) { $existing[$dateKey] = $s }
    }
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("日付,始値,前場引値,後場始値,終値,高値,安値,出来高")
    foreach ($k in ($existing.Keys | Sort-Object)) {
        $r = $existing[$k]
        [void]$lines.Add(("{0},{1},{2},{3},{4},{5},{6},{7}" -f $r.日付, $r.始値, $r.前場引値, $r.後場始値, $r.終値, $r.高値, $r.安値, $r.出来高))
    }
    Set-Content -Path $summaryPath -Value $lines -Encoding UTF8
    Write-Log ("日次サマリー: 全{0}日分に更新" -f ($existing.Keys.Count))

    # 本日(平日)のデータ有無を確認(大引け15:30以降に実行された場合のみ判定)
    $todayJst = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $jst)
    $todayKey = $todayJst.ToString("yyyy-MM-dd")
    if ($todayJst.DayOfWeek -ne "Saturday" -and $todayJst.DayOfWeek -ne "Sunday" -and $todayJst.TimeOfDay -ge [TimeSpan]"15:30:00") {
        if ($byDate1m.ContainsKey($todayKey)) {
            Write-Log ("本日({0})のデータを取得済み" -f $todayKey)
        } else {
            Write-Log ("本日({0})の取引データなし(休場日または取得失敗の可能性)" -f $todayKey)
        }
    }

    if ($byDate1m.Keys.Count -eq 0) {
        Write-Log "異常: 1分足データを1件も取得できませんでした"
        Write-Log "===== 収集終了(異常) ====="
        exit 1
    }
    Write-Log "===== 収集終了(正常) ====="
    exit 0
} catch {
    Write-Log ("致命的エラー: {0}" -f $_.Exception.Message)
    exit 1
}
