# 設定値の読み込みを行う
$configure = @{
    LosslessCut = "C:\Program Files (x86)\LosslessCut-win-x64\"
    ffmpeg = "ffmpeg.exe"
    silence_threshold  = 0.08
    silence_gap        = 2.0
    minimal_silence    = 3.5
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
$exit=$false
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
    exit
}


# 動画選択ダイアログを表示する
[void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
$dialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{Title = "動画ファイルを選択してください"}
$movie = $false
if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
{
    $movie = $dialog.FileName
}
else
{
    exit
}

# 動画があるフォルダに移動する
Set-Location (Split-Path -Path $movie -Parent -Resolve)
# FFmpegで空白検出、時間がかかる
$silence_raw_text=&$configure.ffmpeg -i $movie -af "silencedetect=n=$($configure.silence_threshold):d=1.5:m=0" -vn -f null - 2>&1|% { $_.ToString() }
# 空白の開始と終了を抽出
$starts = echo $silence_raw_text| Select-String "silence_start: (.*)$"  |ForEach-Object { $($_.matches.groups[1]) }|% { [float]$_.ToString() }
$ends = echo $silence_raw_text| Select-String "silence_end: (.*) \|" |ForEach-Object { $($_.matches.groups[1]) }|% { [float]$_.ToString() }

# 短い間があれば連結して長い間にする
$enhanced_starts =New-Object 'System.Collections.Generic.List[float]'
$enhanced_ends   =New-Object 'System.Collections.Generic.List[float]'
$enhanced_starts.Add($starts[0])
$starts=$starts[1..($starts.Length-1)]
for($i = 0; $i -lt $starts.Length; $i++)
{
    if( ($starts[$i] - $ends[$i]) -gt $configure.silence_gap)
    {
        $enhanced_starts.Add($starts[$i])
        $enhanced_ends.Add($ends[$i])
    }
}
$enhanced_ends.Add($ends[-1])
$starts = @($enhanced_starts)
$ends   = @($enhanced_ends)

# 短すぎる間はおかしいので削除する
$enhanced_starts = New-Object 'System.Collections.Generic.List[float]'
$enhanced_ends   = New-Object 'System.Collections.Generic.List[float]'
for($i = 0; $i -lt $starts.Length; $i++)
{
    if( ($ends[$i] - $starts[$i]) -gt $configure.minimal_silence)
    {
        $enhanced_starts.Add($starts[$i])
        $enhanced_ends.Add($ends[$i])
    }
}
$starts = @($enhanced_starts)
$ends   = @($enhanced_ends)

# 反転する
$inverted_starts = New-Object 'System.Collections.Generic.List[float]'
$inverted_ends   = New-Object 'System.Collections.Generic.List[float]'
$inverted_starts.Add(0)
for($i = 0; $i -lt ($starts.Length-1); $i++)
{
        $inverted_starts.Add($ends[$i])
        $inverted_ends.Add($starts[$i])
}
$inverted_ends.Add($starts[-1])

# segment.csvに書き込む
if(Test-Path segment.csv){Remove-Item segment.csv}
for($i = 0; $i -lt (@($inverted_starts).Length); $i++)
{
    [string]$inverted_starts[$i]+","+[string]$inverted_ends[$i]+",seg_"+($i+1) |Out-File -Append -Encoding utf8 segment.csv
}

$sum_length=0
for($i = 0; $i -lt (@($inverted_starts).Length); $i++)
{
    $sum_length = $sum_length + ($inverted_ends[$i] - $inverted_starts[$i])
}
$average_length = $sum_length/$i

if ((@($inverted_starts).Length) -lt 135)
{
    if ($average_length -lt 12)
    {
        [System.Windows.Forms.MessageBox]::Show("完了しましたが、silence_thresholdの調整をお勧めします。`n(もう少し小さく)","info","OK","Information ")
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("完了しましたが、silence_thresholdの調整をお勧めします。`n(もう少し大きく)","info","OK","Information ")
    }
    exit
}

[System.Windows.Forms.MessageBox]::Show("完了しました。","info","OK","Information ")



