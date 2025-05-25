# 設定値の読み込みを行う
$configure = @{
    LosslessCut = "C:\Program Files (x86)\LosslessCut-win-x64\"
    ffmpeg = "ffmpeg.exe"
}
if ( Test-Path "configure.json" )
{
    $configure=Get-Content -Path "configure.json" -Encoding "utf8" | ConvertFrom-Json
}
else
{
    ConvertTo-Json -InputObject $configure | Out-File -Encoding utf8 -FilePath configure.json
}

# ffmpegのパスを特定する
if (!( Test-Path $configure.ffmpeg ))
{
    if (!( test-Path $configure.LosslessCut ))
    {
        Add-Type -AssemblyName System.Windows.Forms
        $result = [System.Windows.Forms.MessageBox]::Show("LosslessCutが見つかりませんでした。`nLossLessCutのフォルダを選択してください。","確認","YesNo","Exclamation","Button2")
        if ($result -eq "Yes")
        {
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ Description = 'LosslessCutフォルダを選択してください'}
            if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $configure.LosslessCut = $dialog.SelectedPath
                $configure.ffmpeg = @(Get-ChildItem -Path $configure.LosslessCut -Recurse -Filter "ffmpeg.exe")[0].FullName
                ConvertTo-Json -InputObject $configure | Out-File -Encoding utf8 -FilePath configure.json
            }
        }
    }
    $configure.ffmpeg = @(Get-ChildItem -Path $configure.LosslessCut -Recurse -Filter "ffmpeg.exe")[0].FullName
    
}
if (!( Test-Path $configure.ffmpeg ))
{
    [System.Windows.Forms.MessageBox]::Show("ffmpegが見つかりませんでした。`n終了します。","エラー","OK","Error")
    exit 1
}
