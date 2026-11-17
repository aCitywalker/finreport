# 启动本机整理页面，并打开一条公网通道。已在运行时不会重复启动。
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root
New-Item -ItemType Directory -Force -Path (Join-Path $root "bin") | Out-Null
$urlFile = Join-Path $root "web\public_url.txt"
$log = Join-Path $root "bin\cloudflared.log"
$exe = Join-Path $root "bin\cloudflared.exe"

$listen = Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue
if (-not $listen) {
    Start-Process -FilePath "py" -ArgumentList "-3", (Join-Path $root "server.py") -WorkingDirectory $root -WindowStyle Hidden
    $ready = $false
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Seconds 1
        $listen = Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue
        if ($listen) { $ready = $true; break }
    }
    if (-not $ready) {
        Write-Output "本机页面未能在 8765 端口启动。"
        exit 1
    }
}

$cf = Get-Process -Name cloudflared -ErrorAction SilentlyContinue
if (-not $cf) {
    if (-not (Test-Path $exe)) {
        Write-Output "缺少 bin\cloudflared.exe，公网地址无法打开。"
        exit 1
    }
    Remove-Item $log -ErrorAction SilentlyContinue
    Start-Process -FilePath $exe -ArgumentList @(
        "tunnel", "--url", "http://127.0.0.1:8765", "--protocol", "http2", "--no-autoupdate", "--logfile", $log
    ) -WorkingDirectory $root -WindowStyle Hidden
    $found = ""
    for ($i = 0; $i -lt 45; $i++) {
        Start-Sleep -Seconds 1
        if (Test-Path $log) {
            $text = Get-Content $log -Raw -ErrorAction SilentlyContinue
            if ($text -match "https://[a-z0-9-]+\.trycloudflare\.com") {
                $found = $Matches[0]
                break
            }
        }
    }
    if (-not $found) {
        Write-Output "公网通道已启动，但还没有读到地址。请稍后查看 bin\cloudflared.log。"
        exit 1
    }
    Set-Content -Path $urlFile -Value $found -Encoding ascii
    Write-Output "公网地址：$found"
} else {
    if (Test-Path $urlFile) {
        Write-Output ("公网地址：" + (Get-Content $urlFile -Raw).Trim())
    } else {
        Write-Output "公网通道已在运行。地址文件不存在时，请看通道日志里的 trycloudflare 链接。"
    }
}
