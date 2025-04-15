param(
    [string]$CsvPath,
    [string]$OutputHtmlPath = "$(Get-Location)\OneDriveQuota_FromCsv.html"
)

if (-not (Test-Path $CsvPath)) {
    Write-Host "CSVファイルが見つかりません: $CsvPath" -ForegroundColor Red
    exit 1
}

$data = Import-Csv -Path $CsvPath -Encoding UTF8
$json = $data | ConvertTo-Json -Compress

# HTMLテンプレート（簡略化版）
$html = @"
<!DOCTYPE html>
<html>
<head>
<title>OneDriveレポート</title>
<script>window.quotaData = $json</script>
</head>
<body>
<!-- コンテンツは簡略化 -->
</body>
</html>
"@

$html | Out-File -FilePath $OutputHtmlPath -Encoding UTF8
Write-Host "HTMLファイルが生成されました: $OutputHtmlPath"
