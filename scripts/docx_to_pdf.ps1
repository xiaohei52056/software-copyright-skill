# docx_to_pdf.ps1 —— Windows 版软著材料转 PDF（LibreOffice）
# 用法: powershell -ExecutionPolicy Bypass -File docx_to_pdf.ps1 正式资料/
#   - 自动探测 soffice：PATH → Program Files → Program Files (x86) → LocalAppData
#   - 只做文件转 PDF，不调用打印机；每个文件最多等 2 分钟，超时杀掉进程并提示
#   - 若弹出"正在等待打印机连接"：LibreOffice 启动时检测默认打印机队列，默认打印机
#     离线/未连接就会弹窗阻塞。转换前请确保默认打印机可用（可临时设为 Microsoft Print to PDF，
#     转完再恢复），或弹窗时点"取消"。
#   - 转换后打印每个 PDF 的大小，超过 50MB 给出警告
# Linux/macOS 请用同目录 docx_to_pdf.sh。

param(
    [string]$Dir = "."
)

$ErrorActionPreference = "Stop"
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8

# ---- 探测 soffice ----
$candidates = @(
    (Get-Command soffice -ErrorAction SilentlyContinue).Source,
    "C:\Program Files\LibreOffice\program\soffice.exe",
    "C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "$env:LOCALAPPDATA\Programs\LibreOffice\program\soffice.exe",
    "$env:ProgramFiles\LibreOffice\program\soffice.exe"
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if (-not $candidates) {
    Write-Host "未找到 LibreOffice（soffice.exe）。请安装后重试，例如：winget install --id TheDocumentFoundation.LibreOffice -e" -ForegroundColor Red
    exit 1
}
Write-Host "使用 soffice: $candidates"

$dir = (Resolve-Path $Dir).Path
$docx = Get-ChildItem -Path $dir -Filter *.docx -File
if (-not $docx) { Write-Host "目录中没有 docx 文件: $dir"; exit 0 }

foreach ($f in $docx) {
    $pdf = Join-Path $dir ($f.BaseName + ".pdf")
    # 目标 PDF 被占用（例如阅读器打开）时先提示
    try {
        $fs = [System.IO.File]::Open($pdf, 'OpenOrCreate', 'ReadWrite', 'None')
        $fs.Close()
    } catch {
        Write-Host ("跳过（PDF 被占用）: " + $f.Name) -ForegroundColor Yellow
        continue
    }
    $proc = Start-Process -FilePath $candidates -ArgumentList @('--headless','--convert-to','pdf','--outdir',$dir,$f.FullName) -PassThru -NoNewWindow
    if (-not $proc.WaitForExit(120000)) {
        Write-Host ("超时（>2 分钟），已停止: " + $f.Name) -ForegroundColor Yellow
        Stop-Process -Id $proc.Id -Force
        continue
    }
    if (Test-Path $pdf) {
        $mb = [math]::Round((Get-Item $pdf).Length / 1MB, 1)
        $warn = if ($mb -gt 50) { "  ⚠️ 超过 50MB" } else { "" }
        Write-Host ("✅ {0}.pdf  {1}MB{2}" -f $f.BaseName, $mb, $warn)
    } else {
        Write-Host ("转换失败: " + $f.Name) -ForegroundColor Red
    }
}
Write-Host "完成。手册目录为 Word 域：LibreOffice 转换时已按 updateFields 自动更新；如目录未刷新，请用 Word 打开手册 F9 更新后另存 PDF。"
