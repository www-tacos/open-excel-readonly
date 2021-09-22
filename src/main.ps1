<#=============================================
Main����
=============================================#>
$ROOT = (Get-Location).Path
$SRC = Join-Path ${ROOT} "src"

# Set-ExcelDefaultOption.psm1 �̓ǂݍ���
Import-Module "${SRC}/Set-ExcelDefaultOption.psm1" -Force -Scope Local

# �N�����[�h�̕ύX
try {
  Clear-Host
  Write-Host "Excel�̃f�t�H���g�����I�����Ă��������B"
  Write-Host ("-" * 50)
  Write-Host "0 : �ҏW�\"
  Write-Host "1 : �ǎ��p"
  Write-Host ("-" * 50)
  while ($True) {
    Write-Host -NoNewLine ">> "
    $mode = Read-Host
    if ($mode -eq '0') {
      # �ҏW�\�ɕύX�i�f�t�H���g����j
      Set-ExcelDefaultOption
      break
    } elseif ($mode -eq '1') {
      # �ǎ��p�ɕύX
      Set-ExcelDefaultOption -IsReadOnly
      break
    } else {
      Write-Host "0�܂���1�őI�����Ă��������B�i�����𒆒f����ꍇ��Ctrl+C�j"
      Write-Host "`n`n"
    }
  }
  Write-Host ""
} catch [System.Security.SecurityException] {
  # �Ǘ��Ҍ������Ȃ������ꍇ
  Write-Error "�Ǘ��Ҍ����Ŏ��s���Ă�������"
} catch {
  # ���̑��G���[
  Write-Error $PSItem.Exception
}

# �I��
Write-Host "`n`n"
$WAITTIME = 5
${WAITTIME}..0 | % {
  Write-Host -NoNewLine "`r�������������܂����B$(($_).ToString().PadLeft($WAITTIME.ToString().Length, ' '))�b��ɃE�C���h�E����܂��B"
  Start-Sleep 1
}
exit 0
