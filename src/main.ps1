<#=============================================
Main処理
=============================================#>
$ROOT = (Get-Location).Path
$SRC = Join-Path ${ROOT} "src"

# Set-ExcelDefaultOption.psm1 の読み込み
Import-Module "${SRC}/Set-ExcelDefaultOption.psm1" -Force -Scope Local

# 起動モードの変更
try {
  Clear-Host
  Write-Host "Excelのデフォルト動作を選択してください。"
  Write-Host ("-" * 50)
  Write-Host "0 : 編集可能"
  Write-Host "1 : 読取専用"
  Write-Host ("-" * 50)
  while ($True) {
    Write-Host -NoNewLine ">> "
    $mode = Read-Host
    if ($mode -eq '0') {
      # 編集可能に変更（デフォルト動作）
      Set-ExcelDefaultOption
      break
    } elseif ($mode -eq '1') {
      # 読取専用に変更
      Set-ExcelDefaultOption -IsReadOnly
      break
    } else {
      Write-Host "0または1で選択してください。（処理を中断する場合はCtrl+C）"
      Write-Host "`n`n"
    }
  }
  Write-Host ""
} catch [System.Security.SecurityException] {
  # 管理者権限がなかった場合
  Write-Error "管理者権限で実行してください"
} catch {
  # その他エラー
  Write-Error $PSItem.Exception
}

# 終了
Write-Host "`n`n"
$WAITTIME = 5
${WAITTIME}..0 | % {
  Write-Host -NoNewLine "`r処理が完了しました。$(($_).ToString().PadLeft($WAITTIME.ToString().Length, ' '))秒後にウインドウを閉じます。"
  Start-Sleep 1
}
exit 0
