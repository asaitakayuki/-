#############################################
# 概要：管理簿からユーザ情報を取得し、サーバ作業を行います。

#############################################


# 環境変数宣言
$file_dir = '' #Ondriveフォルダ
$filename = '' #管理簿ファイル名
$privatekey = '' #秘密鍵
$createUserShellOnAuthServer = '' #サーバ作業実行シェル
$record =  #レコード開始位置
$StartCell =  #セル開始位置
$EndCell =  #セル終了位置
$WaitTimeForOneDrive = 
$templeteiniFile = "”#テンプレート設定ファイル
$iniFile = "" #設定ファイル
$targetUsername = "username" #置換用ユーザ名
$targetPassword = "password" #置換用パスワード名
$iniFileBackupDirectory = "" #設定ファイルバックアップ先

## OneDrive開始
Start-Process -FilePath ""

## エクセル強制終了
$procEx = Get-Process -Name "EXCEL" 2>Out-Null
if ($procEx){
$procEx.Kill()
}

# エクセル開く
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true      # 画面上に表示させる
$excel.DisplayAlerts = $false # 警告メッセージは表示させない

# Dドライブ移動し、作業ディレクトリに移動
Set-Location D:
Set-Location ${file_dir}

# 現在のディレクトリの絶対パスを取得
$currentPath = (Convert-Path .)

# エクセルプロセス待ち
$book = $excel.Workbooks.Open($currentPath + "\" + $filename)

# エクセルファイル閉じる待ち
#Start-Sleep -s 1

# シート名取得
$sheet = $book.Worksheets.Item("sheet1")

# エクセル情報配列化
$input = [PSCustomObject]@{
  bookname  = $filename   # "testdata.xlsx"
  sheetname = $sheet  # "Sheet1"
  startcell = $StartCell  # "D6"
  endcell   = $EndCell    # "W12"
}

# Excelのテーブルをオブジェクトとして取得する関数
Function Get-TableObject($sheet, $start, $end){
  # テーブルオブジェクト作成
  $table = [PSCustomObject]@{
    start = $sheet.Range($start)
    end = $sheet.Range($end)
    key = @()
    data = @()
  }
 
  # テーブルをオブジェクト化
  for($row=$table.start.Row; $row -le $table.end.Row; $row++){
    
    # 1レコード用オブジェクトを準備
    $record = New-Object PSCustomObject
    $key_ref_number = 0
    for($col=$table.start.Column; $col -le $table.end.Column; $col++){
      # 最初にデータの無い列は無視
      if($sheet.cells.item($table.start.Row, $col).text -eq "" ){
        continue
      }
      # 1行目の値からキー名作成
      if($row -eq $table.start.Row){
        $table.key += $sheet.cells.item($row, $col).text
      }
      # 1レコード作成
      else{
        $key_name = ($table.key[$key_ref_number])
        $val = ($sheet.cells.item($row, $col).text)
        $record | Add-Member -MemberType NoteProperty -Name $key_name -Value $val
        $key_ref_number += 1
      }
    }
    # 1レコード追加
    if($row -gt $table.start.Row){
      $table.data += $record
    }
  }
  return $table
  echo $table
}

# Excelのテーブルをオブジェクトとして取得
$table = Get-TableObject $sheet $input.startcell $input.endcell
#return $table.data

function tableDataGet {
    $table.data | %{
        
          # フラグ判定
          if(){
          
            # 未登録ユーザ名およびパスワードを変数に格納
            $unregisteredUsername = $($_.{username})
            $unregisteredPassword = $($_.{password})
            
            ## ユーザ追加
            cmd /C "plink {serverUsername}@{ipaddress} -no-antispoof -i $privatekey $createUserShellOnAuthServer $unregisteredUsername $unregisteredPassword "

            # 複数の戻り値（配列）として、ユーザ情報を返却
            return ${unregisteredUsername}, ${unregisteredPassword}    
         }
    }
}

# ユーザ情報のオブジェクト化
$userInformation = tableDataGet

echo $userInformation

$unRegisterUsername = $userInformation[0] #ユーザ名
$unRegisterPassword = $userInformation[1] #パスワード


# ログイン確認
## iniファイル作成
$(Get-Content $templeteiniFile ) | ForEach-Object {
    $_ -replace "${targetUsername}",$unRegisterUsername `
       -replace "${targetPassword}",$unRegisterPassword
} | Set-Content $iniFile

## Dドライブ移動し、作業ディレクトリに移動
Set-Location D:
Set-Location {workingDirector}

## 接続確認

## プロセス終了待ち
    if ($?) {
        # 接続プロセス戻り待ち
        Start-Sleep -s 5
        # 上書き保存
        #$book.SaveAs($currentPath + "\" + $filename)
        # ブックを閉じる
        $excel.Workbooks.Close()
        #エクセルを閉じる
        $excel.Quit()
		## エクセル強制終了
		$procEx = Get-Process -Name "EXCEL"  2>Out-Null
		if ($procEx){
			$procEx.Kill()
		}       
        # ファイル削除
        del *.opt
        # ディレトリ作成
        New-Item -ItemType directory -Path $iniFileBackupDirectory$unRegisterUsername
        # iniファイル移動
        mv phBuildImage.ini $iniFileBackupDirectory$unRegisterUsername

        # Jenkinsサーバへの成功通知
		return 0
    }else{
        # Jenkinsサーバへの失敗通知
      	exit 1
      	return 1
    }

