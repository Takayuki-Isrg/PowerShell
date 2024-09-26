#------------------------------------------------------
# Windowsサービスをキーワードで検索し、
# 指定したWindowsサービスを起動するスクリプト
# 注意：本スクリプトは管理者権限でないと実行できない
#------------------------------------------------------

# エラーが発生してもスクリプトを停止させない
$ErrorActionPreference = "Continue"

do {
    # Windowsサービス情報を取得
    $services = Get-Service | Select-Object Status, Name, DisplayName

    # 検索ワードを入力
    $searchWord = (Read-Host "検索したいWindowsサービス名を入力してください (*<検索キーワード>*で部分一致検索、Enterキー押下で全量検索します)").ToLower()

    # 検索結果を取得
    $matches = $services | Where-Object { $_.Name.ToLower() -like "*$searchWord*" -or $_.DisplayName.ToLower() -like "*$searchWord*" }

    # 検索結果がヒットしない場合、再度検索するか確認
    if (-not $matches) {
        Write-Host "該当するWindowsサービスが見つかりません。再度検索しますか？ (Yes/No)"
        $confirm = (Read-Host "選択してください (Yes / No)").ToLower()
        if ($confirm -ne "yes") {
            break
        }
        continue
    }

    # スタートアップの種別変更は別ファイルで行う
    # 複数のWindowsサービスがヒットした場合、起動するWindowsサービスを選択
    if ($matches.Count -gt 1) {
        Write-Host "複数のWindowsサービスがヒットしました。起動するWindowsサービスを選択してください"
        for ($i = 0; $i -lt $matches.Count; $i++) {
            Write-Host "$($i + 1): $($matches[$i].DisplayName)"
        }
        # TODO : 「 選択してください 」表示後にEnterキーを押下した時の挙動がおかしい
        $index = [int](Read-Host "選択してください")
        $serviceToStartOrStop = $matches[$index - 1]
    } else {
        $serviceToStartOrStop = $matches[0]
    }

    # 停止状態のWindowsサービスを起動する
    if ($serviceToStartOrStop.Status -eq "Stopped") {
        try {
            Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' の稼働ステータスは '$($serviceToStartOrStop.Status)' です"
            Write-Host "1: 起動"
            Write-Host "2: 再度検索し直す"
            Write-Host "3: ツールの終了"
            $choice = Read-Host "選択してください（1, 2, 3,）"

            Switch($choice) {
                1 {
                    start-Service $serviceToStartOrStop.Name
                    Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' が起動されました"
                }
                2 {
                    Write-Host "検索に戻ります"
                    continue
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
        Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' は既に起動しています"
    }

    # 起動状態のWindowsサービスを停止または再起動する
    if ($serviceToStartOrStop.Status -eq "Running") {
        try {
            Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' の稼働ステータスは '$($serviceToStartOrStop.Status)' です。停止または再起動しますか？"
            Write-Host "1: 停止"
            Write-Host "2: 再起動"
            Write-Host "3: 再度検索する"
            Write-Host "4: ツールの終了"
            $choice = Read-Host "選択してください（1, 2, 3, 4）"

            switch ($choice) {
                1 {
                    Stop-Service $serviceToStartOrStop.Name
                    Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' が停止されました"
                }
                2 {
                    Restart-Service $serviceToStartOrStop.Name
                    Write-Host "Windowsサービス '$($serviceToStartOrStop.DisplayName)' が再起動されました"
                }
                3 {
                    Write-Host "検索に戻ります"
                    continue
                }
                4 {
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
        Write-Host "処理が完了しました"
    }

    # 続けて検索するかの確認
    $confirm = (Read-Host "続けて検索しますか？ (Yes / No)").ToLower()
}
while ($confirm -eq "yes")

Write-Host "ツールを終了します"
exit
