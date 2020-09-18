#############################################
# �T�v�F�Ǘ��납�烆�[�U�����擾���A�T�[�o��Ƃ��s���܂��B

#############################################


# ���ϐ��錾
$file_dir = '' #Ondrive�t�H���_
$filename = '' #�Ǘ���t�@�C����
$privatekey = '' #�閧��
$createUserShellOnAuthServer = '' #�T�[�o��Ǝ��s�V�F��
$record =  #���R�[�h�J�n�ʒu
$StartCell =  #�Z���J�n�ʒu
$EndCell =  #�Z���I���ʒu
$WaitTimeForOneDrive = 
$templeteiniFile = "�h#�e���v���[�g�ݒ�t�@�C��
$iniFile = "" #�ݒ�t�@�C��
$targetUsername = "username" #�u���p���[�U��
$targetPassword = "password" #�u���p�p�X���[�h��
$iniFileBackupDirectory = "" #�ݒ�t�@�C���o�b�N�A�b�v��

## OneDrive�J�n
Start-Process -FilePath ""

## �G�N�Z�������I��
$procEx = Get-Process -Name "EXCEL" 2>Out-Null
if ($procEx){
$procEx.Kill()
}

# �G�N�Z���J��
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true      # ��ʏ�ɕ\��������
$excel.DisplayAlerts = $false # �x�����b�Z�[�W�͕\�������Ȃ�

# D�h���C�u�ړ����A��ƃf�B���N�g���Ɉړ�
Set-Location D:
Set-Location ${file_dir}

# ���݂̃f�B���N�g���̐�΃p�X���擾
$currentPath = (Convert-Path .)

# �G�N�Z���v���Z�X�҂�
$book = $excel.Workbooks.Open($currentPath + "\" + $filename)

# �G�N�Z���t�@�C������҂�
#Start-Sleep -s 1

# �V�[�g���擾
$sheet = $book.Worksheets.Item("sheet1")

# �G�N�Z�����z��
$input = [PSCustomObject]@{
  bookname  = $filename   # "testdata.xlsx"
  sheetname = $sheet  # "Sheet1"
  startcell = $StartCell  # "D6"
  endcell   = $EndCell    # "W12"
}

# Excel�̃e�[�u�����I�u�W�F�N�g�Ƃ��Ď擾����֐�
Function Get-TableObject($sheet, $start, $end){
  # �e�[�u���I�u�W�F�N�g�쐬
  $table = [PSCustomObject]@{
    start = $sheet.Range($start)
    end = $sheet.Range($end)
    key = @()
    data = @()
  }
 
  # �e�[�u�����I�u�W�F�N�g��
  for($row=$table.start.Row; $row -le $table.end.Row; $row++){
    
    # 1���R�[�h�p�I�u�W�F�N�g������
    $record = New-Object PSCustomObject
    $key_ref_number = 0
    for($col=$table.start.Column; $col -le $table.end.Column; $col++){
      # �ŏ��Ƀf�[�^�̖�����͖���
      if($sheet.cells.item($table.start.Row, $col).text -eq "" ){
        continue
      }
      # 1�s�ڂ̒l����L�[���쐬
      if($row -eq $table.start.Row){
        $table.key += $sheet.cells.item($row, $col).text
      }
      # 1���R�[�h�쐬
      else{
        $key_name = ($table.key[$key_ref_number])
        $val = ($sheet.cells.item($row, $col).text)
        $record | Add-Member -MemberType NoteProperty -Name $key_name -Value $val
        $key_ref_number += 1
      }
    }
    # 1���R�[�h�ǉ�
    if($row -gt $table.start.Row){
      $table.data += $record
    }
  }
  return $table
  echo $table
}

# Excel�̃e�[�u�����I�u�W�F�N�g�Ƃ��Ď擾
$table = Get-TableObject $sheet $input.startcell $input.endcell
#return $table.data

function tableDataGet {
    $table.data | %{
        
          # �t���O����
          if(){
          
            # ���o�^���[�U������уp�X���[�h��ϐ��Ɋi�[
            $unregisteredUsername = $($_.{username})
            $unregisteredPassword = $($_.{password})
            
            ## ���[�U�ǉ�
            cmd /C "plink {serverUsername}@{ipaddress} -no-antispoof -i $privatekey $createUserShellOnAuthServer $unregisteredUsername $unregisteredPassword "

            # �����̖߂�l�i�z��j�Ƃ��āA���[�U����ԋp
            return ${unregisteredUsername}, ${unregisteredPassword}    
         }
    }
}

# ���[�U���̃I�u�W�F�N�g��
$userInformation = tableDataGet

echo $userInformation

$unRegisterUsername = $userInformation[0] #���[�U��
$unRegisterPassword = $userInformation[1] #�p�X���[�h


# ���O�C���m�F
## ini�t�@�C���쐬
$(Get-Content $templeteiniFile ) | ForEach-Object {
    $_ -replace "${targetUsername}",$unRegisterUsername `
       -replace "${targetPassword}",$unRegisterPassword
} | Set-Content $iniFile

## D�h���C�u�ړ����A��ƃf�B���N�g���Ɉړ�
Set-Location D:
Set-Location {workingDirector}

## �ڑ��m�F

## �v���Z�X�I���҂�
    if ($?) {
        # �ڑ��v���Z�X�߂�҂�
        Start-Sleep -s 5
        # �㏑���ۑ�
        #$book.SaveAs($currentPath + "\" + $filename)
        # �u�b�N�����
        $excel.Workbooks.Close()
        #�G�N�Z�������
        $excel.Quit()
		## �G�N�Z�������I��
		$procEx = Get-Process -Name "EXCEL"  2>Out-Null
		if ($procEx){
			$procEx.Kill()
		}       
        # �t�@�C���폜
        del *.opt
        # �f�B���g���쐬
        New-Item -ItemType directory -Path $iniFileBackupDirectory$unRegisterUsername
        # ini�t�@�C���ړ�
        mv phBuildImage.ini $iniFileBackupDirectory$unRegisterUsername

        # Jenkins�T�[�o�ւ̐����ʒm
		return 0
    }else{
        # Jenkins�T�[�o�ւ̎��s�ʒm
      	exit 1
      	return 1
    }

