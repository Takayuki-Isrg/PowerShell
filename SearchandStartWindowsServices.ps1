### Windowsサービスをキーワードで検索し、###
### 指定したWindowsサービスを起動するスクリプト ###
## 注意：本スクリプトは管理者権限でないと実行できない ##

# エラーが発生してもスクリプトを停止させない
$ErrorActionPreference = "Continue"

# Windowsサービス情報を取得
$services = Get-Service | Select-Object Status, Name, DisplayName

# 検索ワードを入力
$searchWord = (Read-Host "検索したいWindowsサービス名を入力してください (*<検索キーワード>*で部分一致検索)").ToLower()

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
    $serviceToStartOrStop = $matches[$index - 1]
} else {
    $serviceToStartOrStop = $matches[0]
}

# 停止状態のWindowsサービスを起動する
### start : Fixed on 2024-08-29 続けて検索するかの確認およびループ処理の追加 ###
do {
    if ($serviceToStartOrStop.Status -eq "Stopped") {
        try {
            # Windowsサービスを起動するかどうか確認
            Write-Host `
            "Windowsサービス '$($serviceToStartOrStop.DisplayName)' の稼働ステータスは'$($serviceToStartOrStop.Status)'です" # 2024/08/25追記
            Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' を起動しますか？"
            $YesOrNo = (Read-Host "(Yes / No)").ToLower()

            # 'Yes'が入力された時の処理
            if ($YesOrNo.ToLower() -like "*Yes*") {
                Start-Service $serviceToStartOrStop.Name
                Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' が起動されました"
                # 続けて検索するかを確認
                Write-Host "続けて検索しますか？"
                $confirm = (Read-Host "(Yes / No)").ToLower()
            } else {
                Write-Host "ツールを終了します"
                exit # 2024/08/25追記
            }
        } catch {
            Write-Host "エラーが発生しました: $($_.Exception.Message)"
            Write-Warning "スタックトレース:"
            $_.Exception.StackTrace | Out-String
        }
    } else {
        # TODO : 63行目の文言の修正が必要
        Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' は既に起動しています"
        Write-Host "無効な選択です。ツールを終了します"
    }
}
while ($confirm.ToLower() -like "*Yes*")
### end : Fixed on 2024-08-29 続けて検索するかの確認およびループ処理の追加 ###

# 起動状態のWindowsサービスを停止または再起動する
### start : Fixed on 2024-08-29 続けて検索するかの確認およびループ処理の追加 ###
do {
    if ($serviceToStartOrStop.Status -eq "Running") {
        try {
            ### start : Fixed on 2024-08-26 ###
            # Windowsサービスを停止または再起動するかどうか確認
            Write-Host `
            "Windowsサービス '$($serviceToStartOrStop.DisplayName)' の稼働ステータスは '$($serviceToStartOrStop.Status)' です。停止または再起動しますか？"
            Write-Host "1: 停止"
            Write-Host "2: 再起動"
            Write-Host "3: 終了"
            $choice = Read-Host "選択してください（1, 2, 3）"

            switch ($choice) {
                1 {
                    Stop-Service $serviceToStartOrStop.Name
                    Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' が停止されました"
                    # 続けて検索するかを確認
                    Write-Host "続けて検索しますか？"
                    $confirm = (Read-Host "(Yes / No)").ToLower()
                }
                2 {
                    Restart-Service $serviceToStartOrStop.Name
                    Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' が再起動されました"
                    # 続けて検索するかを確認
                    Write-Host "続けて検索しますか？"
                    $confirm = (Read-Host "(Yes / No)").ToLower()
                }
                3 {
                    Write-Host "ツールを終了します"
                    exit
                }
                default {
                    Write-Host "無効な選択です。ツールを終了します"
                    exit
                }
            }

        } catch {
            Write-Host "エラーが発生しました: $($_.Exception.Message)"
            Write-Warning "スタックトレース:"
            $_.Exception.StackTrace | Out-String
        }
    } else {
        # TODO : 117行目の文言の修正が必要
        Write-Host "無効な選択です。ツールを終了します"
        exit
    ### end : Fixed on 2024-08-26 ###
    }
}
while ($confirm.ToLower() -like "*Yes*")
### end : Fixed on 2024-08-29 続けて検索するかの確認およびループ処理の追加 ###
