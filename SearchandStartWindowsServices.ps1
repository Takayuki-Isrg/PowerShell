### Windowsサービスをキーワードで検索し、###
### 指定したWindowsサービスを起動するスクリプト ###
## 注意：本スクリプトは管理者権限でないと実行できない ##

# エラーが発生してもスクリプトを停止させない
$ErrorActionPreference = "Continue"

# Windowsサービス情報を取得
$services = Get-Service | Select-Object Status, Name, DisplayName

# 検索ワードを入力
$searchWord = (Read-Host "検索したいWindowsサービス名を入力してください (<検索キーワード>*で部分一致検索)").ToLower()

# 検索結果を取得
$matches = $services | Where-Object { $_.Name.ToLower() -like $searchWord -or $_.DisplayName.ToLower() -like $searchWord }

# 検索結果がヒットしない場合の処理
if (-not $matches) {
    Write-Host "該当するWindowsサービスが見つかりません"
    exit
}

# 複数のWindowsサービスがヒットした場合、起動するWindowsサービスを選択
if ($matches.Count -gt 1) {
    Write-Host "複数のWindowsサービスがヒットしました。起動するWindowsサービスを選択してください"
    for ($i = 0; $i -lt $matches.Count; $i++) {
        Write-Host "$($i + 1): $($matches[$i].DisplayName)"
    }
    $index = [int](Read-Host "選択してください")
    $serviceToStart = $matches[$index - 1]
} else {
    $serviceToStart = $matches[0]
}

# 停止状態のWindowsサービスを起動する
if ($serviceToStart.Status -eq "Stopped") {
    try {
        # Windowsサービスを起動するかどうか確認
        Write-Host "Windowsサービス '$($serviceToStart.DisplayName)' を起動しますか？"
        $YesOrNo = (Read-Host "(Yes / No)").ToLower()

        # 'Yes'が入力された時の判定を行う
        if ($YesOrNo.ToLower() -like "*Yes*") {
            Start-Service $serviceToStart.Name
            Write-Host "Windowsサービス '$($serviceToStart.DisplayName)' が起動されました"
        } else {
            Write-Host "ツールを終了します"
        }
    } catch {
        Write-Host "エラーが発生しました: $($_.Exception.Message)"
        Write-Warning "スタックトレース:"
        $_.Exception.StackTrace | Out-String
    }
} else {
    Write-Host "Windowsサービス '$($serviceToStart.DisplayName)' は既に起動しています"
}

# 起動状態のWindowsサービスを停止する
if ($serviceToStart.Status -eq "Running") {
    try {
        # Windowsサービスを停止するかどうか確認
        Write-Host "Windowsサービス '$($serviceToStart.DisplayName)' を停止しますか？"
        $YesOrNo = (Read-Host "(Yes / No)").ToLower()

        # 'Yes'が入力された時の判定を行う
        if ($YesOrNo.ToLower() -like "*Yes*") {
            Start-Service $serviceToStart.Name
            Write-Host "Windowsサービス '$($serviceToStart.DisplayName)' が停止されました"
        } else {
            Write-Host "ツールを終了します"
        }
    } catch {
        Write-Host "エラーが発生しました: $($_.Exception.Message)"
        Write-Warning "スタックトレース:"
        $_.Exception.StackTrace | Out-String
    }
} else {
    Write-Host "Windowsサービス '$($serviceToStart.DisplayName)' は既に停止しています"
}
