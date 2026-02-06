param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Files
)

# CSV が渡されていない場合
if (-not $Files -or $Files.Count -eq 0) {
    Write-Host "CSVファイルをこのスクリプトにドラッグ＆ドロップしてください。" -ForegroundColor Yellow
    Read-Host "Press Enter to finish"
    exit
}

foreach ($file in $Files) {

    # 存在チェック
    if (-not (Test-Path -LiteralPath $file)) {
        Write-Host "File not found: $file" -ForegroundColor Red
        continue
    }

    # 拡張子チェック
    if ([IO.Path]::GetExtension($file).ToLower() -ne ".csv") {
        Write-Host "Not a CSV file: $file" -ForegroundColor Red
        continue
    }

    try {
        # 🔹 読み取り専用（変更なし）
        # Excel由来なら Encoding Default が無難
        $csv = Import-Csv -LiteralPath $file -Encoding Default
        Write-Host "CSV loading completed" -ForegroundColor Cyan
        # 列名を表示（確認用）
        if ($csv.Count -gt 0) {
            $columns = $csv[0].PSObject.Properties.Name
            #Write-Host " 列: $($columns -join ', ')"
        }
    }
    catch {
        Write-Host "読み取りエラー: $file" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
}





# powershell -File cdp.ps1

# Edgeのパスを取得する
function Get-EdgePath {
    $paths = @()

    $pf = [Environment]::GetEnvironmentVariable('ProgramFiles')
    if ($pf) {
        $paths += (Join-Path $pf 'Microsoft\Edge\Application\msedge.exe')
    }

    $pf86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if ($pf86) {
        $paths += (Join-Path $pf86 'Microsoft\Edge\Application\msedge.exe')
    }

    foreach ($p in $paths) {
        if (Test-Path $p) {
            return $p
        }
    }

    $cmd = Get-Command 'msedge.exe' -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Path) {
        return $cmd.Source
    }

    throw "Microsoft Edge (msedge.exe) が見つかりませんでした。"
}

Add-Type -AssemblyName System.Net.Http
# PowerShell5.1だと不要
# Add-Type -AssemblyName System.Net.WebSockets

$global:cdpSocket = [System.Net.WebSockets.ClientWebSocket]::new()
$global:cdpCts    = New-Object System.Threading.CancellationTokenSource
$global:cdpId     = 0

# CDPコマンドを送信
function Send-CdpCommand {
    param(
        [int]$Id,
        [string]$Method,
        [hashtable]$Params
    )

    if (-not $Params) { $Params = @{} }

    $payload = @{
        id     = $Id
        method = $Method
        params = $Params
    } | ConvertTo-Json -Depth 5

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    $seg   = New-Object System.ArraySegment[byte] (, $bytes)

    $global:cdpSocket.SendAsync(
        $seg,
        [System.Net.WebSockets.WebSocketMessageType]::Text,
        $true,
        $global:cdpCts.Token
    ).Wait()
}

# CDPからのメッセージを受信
function Receive-CdpMessageForId {
    param([int]$ExpectedId)

    # 1 回のメッセージを受けるバッファ
    $buffer = New-Object byte[] 8192

    while ($true) {
        # ← ここがポイント。ArraySegment[byte] の作り方
        $segment = New-Object "System.ArraySegment[byte]" -ArgumentList (, $buffer)

        $result = $global:cdpSocket.ReceiveAsync(
            $segment,
            $global:cdpCts.Token
        ).Result

        $json = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count)

        # たまに通知イベントなどで壊れた/期待しない JSON が来るかもしれないので try-catch
        try {
            $msg = $json | ConvertFrom-Json
        } catch {
            continue
        }

        if ($msg.id -eq $ExpectedId) {
            return $msg
        }

        # id が違う通知 (Page.loadEventFired など) は無視して次ループ
    }
}


# CDPにコマンドを送信して受信を待つ
function Invoke-Cdp {
    param(
        [string]$Method,
        [hashtable]$Params
    )

    $global:cdpId++
    $id = $global:cdpId
    Send-CdpCommand -Id $id -Method $Method -Params $Params
    return (Receive-CdpMessageForId -ExpectedId $id)
}

# CDPに接続
function Connect-Cdp {
    $targets = Invoke-RestMethod http://localhost:9222/json/list
    $target  = $targets[0]
    $wsUrl   = $target.webSocketDebuggerUrl

    $uri = [Uri]$wsUrl
    $global:cdpSocket.ConnectAsync($uri, $global:cdpCts.Token).Wait()

    $global:cdpId = 0
    Invoke-Cdp -Method "Page.enable" -Params @{}
    Invoke-Cdp -Method "DOM.enable"  -Params @{}
}



# 背景色を読む（推奨パス）
function Get-BackgroundColorBySelector {
    param([string]$Selector)
    $sel = $Selector.Replace('\','\\').Replace('"','\"')
    $js = @"
(() => {
  const el = document.querySelector("$sel");
  if (!el) return { ok:false, reason: "notfound" };
  const cs = window.getComputedStyle(el);
  return { ok:true, value: cs.backgroundColor };
})()
"@
    $r = Invoke-Cdp -Method "Runtime.evaluate" -Params @{ expression = $js; returnByValue = $true }
    if (-not $r.result.result.value.ok) { throw "Selector '$Selector' が見つかりませんでした。" }
    return $r.result.result.value.value
}

function Convert-RgbToHex {
    param([string]$Rgb)
    if ($Rgb -match 'rgba?\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)') {
        return ('#{0:X2}{1:X2}{2:X2}' -f [int]$Matches[1],[int]$Matches[2],[int]$Matches[3])
    } else { throw "未対応の形式: $Rgb" }
}











# 要素の中央の座標を取得
function Get-ElementCenter {
    param($model)

    $xs = @()
    $ys = @()

    for ($i = 0; $i -lt $model.content.Count; $i += 2) {
        $xs += [double]$model.content[$i]
        $ys += [double]$model.content[$i + 1]
    }

    $cx = ($xs | Measure-Object -Average).Average
    $cy = ($ys | Measure-Object -Average).Average

    return @{ X = $cx; Y = $cy }
}

# 指定の要素をクリックする
function Invoke-CdpClickBySelector {
    param(
        [string]$Selector
    )
    # clickイベントについてはpyppeteerの以下参考
    # https://github.com/pyppeteer/pyppeteer/blob/7dc91ee5173d3836f77800a3774beeaf2b448c0e/pyppeteer/input.py#L285
    $doc    = Invoke-Cdp -Method "DOM.getDocument" -Params @{}
    $rootId = $doc.result.root.nodeId

    $q = Invoke-Cdp -Method "DOM.querySelector" -Params @{
        nodeId  = $rootId
        selector = $Selector
    }

    $nodeId = $q.result.nodeId
    if (-not $nodeId) {
        throw "Selector '$Selector' に一致する要素がありません。"
    }

    $box    = Invoke-Cdp -Method "DOM.getBoxModel" -Params @{ nodeId = $nodeId }
    $center = Get-ElementCenter -model $box.result.model

    Invoke-Cdp -Method "Input.dispatchMouseEvent" -Params @{
        type       = "mousePressed"
        x          = $center.X
        y          = $center.Y
        button     = "left"
        clickCount = 1
    }

    Invoke-Cdp -Method "Input.dispatchMouseEvent" -Params @{
        type       = "mouseReleased"
        x          = $center.X
        y          = $center.Y
        button     = "left"
        clickCount = 1
    }
}

# 指定の要素が表示されるまで待機
function Wait-CdpElementBySelector {
    param(
        [string]$Selector,
        [int]$TimeoutMs  = 10000,
        [int]$IntervalMs = 250
    )

    $elapsed = 0
    while ($elapsed -lt $TimeoutMs) {
        $doc    = Invoke-Cdp -Method "DOM.getDocument" -Params @{}
        $rootId = $doc.result.root.nodeId

        $q = Invoke-Cdp -Method "DOM.querySelector" -Params @{
            nodeId  = $rootId
            selector = $Selector
        }

        if ($q.result.nodeId) {
            return $true
        }

        Start-Sleep -Milliseconds $IntervalMs
        $elapsed += $IntervalMs
    }

    throw "Timeout: Selector '$Selector' が見つかりませんでした。"
}

# 指定の要素にテキストを挿入
function Send-CdpText {
    param(
        [string]$Text
    )

    # フォーカスされた要素にテキストを挿入
    Invoke-Cdp -Method "Input.insertText" -Params @{ text = $Text }
}

# Enter押下
function Send-CdpEnter {
    # keyDown
    Invoke-Cdp -Method "Input.dispatchKeyEvent" -Params @{
        type                  = "keyDown"
        key                   = "Enter"
        code                  = "Enter"
        windowsVirtualKeyCode = 13
        nativeVirtualKeyCode  = 13
        text = "`r"
    }

    # keyUp
    Invoke-Cdp -Method "Input.dispatchKeyEvent" -Params @{
        type                  = "keyUp"
        key                   = "Enter"
        code                  = "Enter"
        windowsVirtualKeyCode = 13
        nativeVirtualKeyCode  = 13
        text = "`r"
    }
}

# 特定のテキストが表示されるまで待機
function Wait-CdpText {
    param(
        [string]$Text,
        [int]$TimeoutMs = 10000,
        [int]$IntervalMs = 300
    )

    $elapsed = 0

    while ($elapsed -lt $TimeoutMs) {

        # JS で文字列を検索して返す...入力を信用しないならエスケープしたほうがいい
        $js = @"
document.body.innerText.includes("$Text")
"@

        $r = Invoke-Cdp -Method "Runtime.evaluate" -Params @{
            expression    = $js
            returnByValue = $true
        }

        if ($r.result.result.value -eq $true) {
            #Write-Host "Text '$Text' detected."
            return $true
        }

        Start-Sleep -Milliseconds $IntervalMs
        $elapsed += $IntervalMs
    }

    throw "Timeout waiting for text '$Text'"
}


# 自動操作用のEdgeを起動
$edgePath = Get-EdgePath
$userDataDir = Join-Path $env:TEMP "edge-devtools-profile"
if (-not (Test-Path $userDataDir)) {
    New-Item -ItemType Directory -Path $userDataDir | Out-Null
}

$edge = Start-Process -FilePath $edgePath -ArgumentList @(
    "--remote-debugging-port=9222",
    "--user-data-dir=$userDataDir",
    "--no-first-run",
    "--new-window"
) -PassThru

Start-Sleep -Seconds 2

# CDPを接続
Connect-Cdp | Out-Null


# STEP 1: navigate https://zenn.dev/
Invoke-Cdp -Method "Page.navigate" -Params @{ url = "http://10.136.24.12/PanaCIMMC/App/LogOn" } | Out-Null

# 表示を待つ
#Wait-CdpElementBySelector -Selector "#txtLoginId"
Wait-CdpText "Supervisor" | Out-Null
Start-Sleep -Milliseconds 300

# STEP 4: click 入力フォーム
Invoke-CdpClickBySelector -Selector "#txtLoginId" | Out-Null
Start-Sleep -Milliseconds 200

# STEP 5: change value -> "1"
Send-CdpText -Text "1"
Send-CdpEnter


#Wait-CdpElementBySelector -Selector "crudForm"

Wait-CdpText "Inventory" | Out-Null
Start-Sleep -Milliseconds 200
Invoke-CdpClickBySelector -Selector "#btnReceiveMaterial" | Out-Null
Wait-CdpText "Create a Receive Material" | Out-Null
Start-Sleep -Milliseconds 200
Invoke-CdpClickBySelector -Selector "#chkAutoRegister" | Out-Null
Start-Sleep -Milliseconds 100


#CSVを繰り返す

foreach ($row in $csv) {
    $QR = $row.OKQR

#$QR = "VW2-3656-103;MID]1A2502040449;Q]10000;ADS]7A21;VND]R354;LOT]PL4UT1500710010000;MFG]20260127;EXP]20310127"
Invoke-CdpClickBySelector -Selector "#txtHomeScan" | Out-Null
Start-Sleep -Milliseconds 100 | Out-Null
Send-CdpText -Text $QR | Out-Null
Send-CdpEnter | Out-Null
Wait-CdpElementBySelector -Selector "#notificationBanner" | Out-Null

$bg = Get-BackgroundColorBySelector -Selector "#notificationBanner"
$hex = Convert-RgbToHex $bg
#Write-Host "notificationBanner background-color = $bg ($hex)"
#E53935→NG Red
#43A047→OK Green
$expected = "#43A047"
if ($hex.ToUpper() -eq $expected) {
    Write-Host "Receive success" -ForegroundColor Green
    Write-Host $QR -ForegroundColor Green
} else {
    Write-Host "Receive failure" -ForegroundColor Red
    Write-Host $QR -ForegroundColor Red
}
}




# CDP WebSocket を閉じる
$global:cdpSocket.CloseAsync(
    [System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure,
    "done",
    $global:cdpCts.Token
).Wait()

$global:cdpSocket.Dispose()
$global:cdpCts.Dispose()

if ($edge -and !$edge.HasExited) {
    $edge.Kill()
}
Read-Host "Press Enter to finish"