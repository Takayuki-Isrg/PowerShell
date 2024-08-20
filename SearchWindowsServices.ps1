# �G���[���������Ă��X�N���v�g���~�����Ȃ�
$ErrorActionPreference = "Continue"

# �T�[�r�X�����擾
$services = Get-Service | Select-Object Status, Name, DisplayName

# �������[�h�����
$searchWord = (Read-Host "�����������T�[�r�X������͂��Ă������� (���C���h�J�[�h��)").ToLower()

# �������ʂ��擾
$matches = $services | Where-Object { $_.Name.ToLower() -like $searchWord -or $_.DisplayName.ToLower() -like $searchWord }

# �������ʂ��Ȃ��ꍇ
if (-not $matches) {
    Write-Host "�Y������T�[�r�X��������܂���"
    exit
}

# �����̃T�[�r�X���q�b�g�����ꍇ�A�N������T�[�r�X��I��
if ($matches.Count -gt 1) {
    Write-Host "�����̃T�[�r�X���q�b�g���܂����B�N������T�[�r�X��I�����Ă�������"
    for ($i = 0; $i -lt $matches.Count; $i++) {
        Write-Host "$($i + 1): $($matches[$i].DisplayName)"
    }
    $index = Read-Host "�I�����Ă�������:" -as [int]
    $serviceToStart = $matches[$index - 1]
} else {
    $serviceToStart = $matches[0]
}

# �I�������T�[�r�X����~��Ԃ̏ꍇ�ɋN��
if ($serviceToStart.Status -eq "Stopped") {
    try {
		    Start-Service $serviceToStart.Name
		    Write-Host "�T�[�r�X '$($serviceToStart.DisplayName)' ���N������܂���"
    } catch {
		    Write-Host "�G���[���������܂���: $($_.Exception.Message)"
		    Write-Warning "�X�^�b�N�g���[�X:"
		    $_.Exception.StackTrace | Out-String
}
} else {
    Write-Host "�T�[�r�X '$($serviceToStart.DisplayName)' �͊��ɋN�����Ă��܂�"
}