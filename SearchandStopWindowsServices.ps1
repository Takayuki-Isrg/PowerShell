# エラーが発生してもスクリプトを停止させない
$ErrorActionPreference = "Continue"

# サービス情報を取得
$services = Get-Service | Select-Object Status, Name, DisplayName

# 検索ワードを入力
$searchWord = (Read-Host "検索したいサービス名を入力してください (ワイルドカード可)").ToLower()

# 検索結果を取得
$matches = $services | Where-Object { $_.Name.ToLower() -like $searchWord -or $_.DisplayName.ToLower() -like $searchWord }

# 検索結果がない場合
if (-not $matches) {
    Write-Host "該当するサービスが見つかりません"
    exit
}

# 複数のサービスがヒットした場合、停止するサービスを選択
if ($matches.Count -gt 1) {
    Write-Host "複数のサービスがヒットしました。停止するサービスを選択してください"
    for ($i = 0; $i -lt $matches.Count; $i++) {
        Write-Host "$($i + 1): $($matches[$i].DisplayName)"
    }
    $index = Read-Host "選択してください:" -as [int]
    $serviceToStop = $matches[$index - 1]
} else {
    $serviceToStop = $matches[0]
}

# 選択したサービスが起動状態の場合に停止
if ($serviceToStop.Status -eq "Stopped") {
    try {
		    Stop-Service $serviceToStop.Name
		    Write-Host "サービス '$($serviceToStop.DisplayName)' が停止されました"
    } catch {
		    Write-Host "エラーが発生しました: $($_.Exception.Message)"
		    Write-Warning "スタックトレース:"
		    $_.Exception.StackTrace | Out-String
}
} else {
    Write-Host "サービス '$($serviceToStop.DisplayName)' は既に停止しています"
}