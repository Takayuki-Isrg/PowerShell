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

# 複数のサービスがヒットした場合、起動するサービスを選択
if ($matches.Count -gt 1) {
  Write-Host "複数のサービスがヒットしました。起動するサービスを選択してください"
  for ($i = 0; $i -lt $matches.Count; $i++) {
    Write-Host "$($i+1): $($matches[$i].DisplayName)"
  }
  $index = Read-Host "選択してください:" -as [int]
  $serviceToStart = $matches[$index - 1]
} else {
    $serviceToStart = $matches[0]
}

# 選択したサービスが停止状態の場合に起動
if ($serviceToStart.Status -eq "Stopped") {
		try {
				Start-Service $serviceToStart.Name
				Write-Host "サービス '$($serviceToStart.DisplayName)' が起動されました"
		} catch {
				Write-Host "エラーが発生しました: $($_.Exception.Message)"
				Write-Warning "スタックトレース:"
				$_.Exception.StackTrace | Out-String
		}
} else {
		Write-Host "サービス '$($serviceToStart.DisplayName)' は既に起動しています"
}