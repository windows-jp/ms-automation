### �t�@�C���쐬/�ҏW ###
# ��̃t�@�C���쐬
New-Item file.txt

# �R�}���h�̌��ʂ��t�@�C���ɏo�� (���{��OK)
Get-ChildItem C:\Users\opkd0\Desktop | Out-File desktop.txt


### CSV ### 
# �R�}���h�̌��ʂ�csv�t�@�C���ɏo�� (���{��OK: Unicode -> SJIS)
Get-ChildItem C:\Users\opkd0\Desktop | Export-Csv desktop.csv -NoType -Encoding Default

# csv�t�@�C����ǂݍ��� (���{��OK: SJIS -> Unicode)
$csv = Import-Csv .\desktop.csv -Encoding Default

# �J���������w�肵��csv�t�@�C���ɏo�� (���{��OK: Unicode -> SJIS)
$csv | Select-Object Name, LastWriteTime | Export-Csv desktop_filenames.csv -NoType -Encoding Default

# �w�b�_�[��񂪂Ȃ�csv�t�@�C����ǂݍ���
$csv = Get-Content .\sample.csv | ConvertFrom-Csv -Header Name,Date


### �ҏW ###
# �������u�����ď㏑���ۑ�
(Get-Content '.\sample3.csv').Replace('"','').Replace('-','') | Set-Content .\sample3.csv

# �擪�s���폜���ď㏑���ۑ�
$content = Get-Content .\sample3.csv
$content[0] = $null # �擪�s = 0
$content | Set-Content .\sample3.csv

# �d�����폜���ď㏑���ۑ�
Get-Content .\sample4.txt | Sort-Object | Get-Unique | Set-Content .\sample4.txt


### �t�@�C�����擾 ###
# �s���擾
(Get-Content '.\sample3.csv').Length

# grep (���K�\���K�p & �啶������������ʂ���)
$pattern = "^aM.*?:"
(Select-String $pattern .\sample3.csv -CaseSensitive -Encoding Default).Line # -> String

# grep -v (���K�\���K�p & �啶������������ʂ���)
$pattern = "^aM.*?:"
(Select-String $pattern .\sample3.csv -NotMatch -CaseSensitive -Encoding Default).Line # -> String

# ���K�\���Ƀ}�b�`�������������𒊏o [sed s/(^aM.*?)/\1/g�Ɠ���]
$pattern = "^aM.*?:"
(Select-String "^aM.*?:" .\sample3.csv -CaseSensitive -Encoding Default).Matches.Value


### ���p�� ###
# �p�^�[���Ƀ}�b�`��������������u�����ď㏑���ۑ�
$pattern = "[0-9]{8}"
$before = (Select-String $pattern .\sample.sql).Matches.Value | Get-Unique
$after = '20220401'
(Get-Content .\sample.sql).Replace($before, $after) | Set-Content .\sample.sql