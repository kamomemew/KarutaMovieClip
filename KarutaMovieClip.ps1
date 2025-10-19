Import-Module ".\KarutaMovieClip.psm1"
# 設定値の読み込みを行う
$configure = @{
    LosslessCut       = "C:\Program Files (x86)\LosslessCut-win-x64\"
    ffmpeg            = "ffmpeg.exe"
    silence_threshold = 0.67
}
# configure.jsonの存在確認
Write-Host -NoNewline "Checking config file ... "
if ( Test-Path "configure.json" ) {
    Write-Host -ForegroundColor Green "[OK]"
    $configure = Get-Content -Path "configure.json" -Encoding "utf8" | ConvertFrom-Json
}
else {
    Write-Host -ForegroundColor Red "[NG]"
    ConvertTo-Json -InputObject $configure | Out-File -Encoding utf8 -FilePath configure.json
}

# ffmpegのパスを特定する

if (-not ( Test-Path $configure.ffmpeg )) {
    Write-Host -NoNewline "LosslessCut installed ... "
    if (-not ( test-Path $configure.LosslessCut )) {
        Write-Host -ForegroundColor Red "[NG]"
        Add-Type -AssemblyName System.Windows.Forms
        $result = [System.Windows.Forms.MessageBox]::Show("LosslessCutが見つかりませんでした。`nLossLessCutのフォルダを選択してください。", "確認", "YesNo", "Exclamation", "Button2")
        if ($result -eq "Yes") {
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ Description = 'LosslessCutフォルダを選択してください' }
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $configure.LosslessCut = $dialog.SelectedPath
                $configure.ffmpeg = @(Get-ChildItem -Path $configure.LosslessCut -Recurse -Filter "ffmpeg.exe")[0].FullName
            }
        }
        else{
            exit
        }
    }
    else{
        Write-Host -ForegroundColor Green "[OK]"
    }
    $configure.ffmpeg = @(Get-ChildItem -Path $configure.LosslessCut -Recurse -Filter "ffmpeg.exe")[0].FullName
    ConvertTo-Json -InputObject $configure | Out-File -Encoding utf8 -FilePath configure.json
}
Write-Host -NoNewline "FFmpeg installed ... "
if (-not ( Test-Path $configure.ffmpeg )) {
    Write-Host -ForegroundColor Red "[NG]"
    [System.Windows.Forms.MessageBox]::Show("ffmpegが見つかりませんでした。`n終了します。", "エラー", "OK", "Error")
    exit 1
}
Write-Host -ForegroundColor Green "[OK]"

# 動画選択ダイアログを表示する
[void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
$dialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{Title = "動画ファイルを選択してください"}
$movie = $false
if($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK){
    exit 1
}
$movie = $dialog.FileName

# 動画があるフォルダに移動する
Set-Location (Split-Path -Path $movie -Parent -Resolve)
# 音量スレッショルド用の値取得
Write-Host -NoNewline "Volume threshold setting ... "
if($configure.silence_threshold | Select-String "dB" -Quiet){
    $silence_threshold=([float]($configure.silence_threshold.Replace("dB",""))).ToString("0.00")
}
else {
    $volume_raw_text=&$configure.ffmpeg -v error -ss 600  -t 360 -i $movie -af "firequalizer=gain='if(lt(f,1600),if(gt(f,350),0,-INF),-INF)',aresample=44100,asetnsamples=2205,astats=reset=1:metadata=1,ametadata=print:key=lavfi.astats.Overall.peak_level:file='pipe\:1'" -vn -f null - |& { process{ $_.ToString() }}
    $peaks=$volume_raw_text| select-String "peak_level=(.*)$" | &{ process {[float]$($_.matches.groups[1]).ToString() }}
    $silence_threshold=(Get-Percentile -Sequence $peaks -Percentile $configure.silence_threshold).ToString("0.00")
}
Write-Host -ForegroundColor Green "$silence_threshold dB"
Write-Host -NoNewline "Silence detection ... "
# FFmpegで空白検出、時間がかかる
$silence_raw_text=&$configure.ffmpeg -i $movie -af "firequalizer=gain='if(lt(f,1600),if(gt(f,350),0,-INF),-INF)',silencedetect=n=$($silence_threshold)dB:d=0.8:m=0" -vn -f null - 2>&1|ForEach-Object { $_.ToString() }
Write-Host -ForegroundColor Green "[OK]"
# 空白の開始と終了を抽出
$starts_Array =  $silence_raw_text| Select-String "silence_start: (.*)$"  |ForEach-Object { [float]$($_.matches.groups[1]).ToString() }
$ends_Array = $silence_raw_text| Select-String "silence_end: (.*) \|" |ForEach-Object { [float]$($_.matches.groups[1]).ToString() }
$starts = New-Object System.Collections.Generic.List[float]
$ends   = New-Object System.Collections.Generic.List[float]
$starts.AddRange([float[]]$starts_Array);$ends.AddRange([float[]]$ends_Array)

# 短い間があれば連結して長い間にする
$starts,$ends = Merge-Gaps -starts $starts -ends $ends -sec 3.9
# あまりに短い間はおかしいので削除する
#$starts,$ends = Select-Span -starts $starts -ends $ends -gt 1.4
# 反転し、「有音部」を抜き出す
$starts,$ends = Select-Invert -starts $starts -ends $ends
#$starts,$ends = Select-Span -starts $starts -ends $ends -gt 2
$starts,$ends = Deny-Gap -before -after -starts $starts -ends $ends -sec 7
#読みのとき。
#$starts,$ends = Merge-Gaps -starts $starts -ends $ends -sec 4

if ($starts.Count -lt 135)
{
    $lengths = 0..($starts.Count - 1 ) | ForEach-Object { $ends[$_] - $starts[$_] }
    $average_length = ($lengths | Measure-Object -Average).Average
    if ($average_length -lt 12){
        [System.Windows.Forms.MessageBox]::Show("完了しましたが、silence_thresholdの調整をお勧めします。`n(もう少し小さく)","info","OK","Information ")
    }
    else{
        [System.Windows.Forms.MessageBox]::Show("完了しましたが、silence_thresholdの調整をお勧めします。`n(もう少し大きく)","info","OK","Information ")
    }
}

$proceed=[System.Windows.Forms.MessageBox]::Show("続けて上の句検出しますか？`n(実験的機能です。)","info","YesNo","Information ")
if($proceed -eq "No"){
    if(Test-Path segment.csv){Remove-Item segment.csv}
    for($i = 0; $i -lt $starts.Count; $i++){
        [string]$starts[$i]+","+[string]$ends[$i]+",seg_"+($i+1) |Out-File -Append -Encoding utf8 segment.csv
    }
    exit 0
}

if(Test-Path segment.csv){Remove-Item segment.csv}
$kamishimo = Get-KamiShimoList -starts $starts -ends $ends
$c=0
for($i = 0; $i -lt $starts.Count; $i++){
    if ($kamishimo[$i] | Select-String "kami" -Quiet){
        [string]$starts[$i]+","+[string]$ends[$i]+","+[string]$kamishimo[$i]+($c+1) |Out-File -Append -Encoding utf8 segment.csv
        $c=$c+1
    }
}
$c=0
for($i = 0; $i -lt $starts.Count; $i++){
    if ($kamishimo[$i] | Select-String "shimo" -Quiet){
        [string]$starts[$i]+","+[string]$ends[$i]+","+[string]$kamishimo[$i]+($c+1) |Out-File -Append -Encoding utf8 segment.csv
        $c=$c+1
    }
}
$c=0
for($i = 0; $i -lt $starts.Count; $i++){
    
    if ($kamishimo[$i] | Select-String "\?" -Quiet){
        [string]$starts[$i]+","+[string]$ends[$i]+","+[string]$kamishimo[$i]+($c+1) |Out-File -Append -Encoding utf8 segment.csv
        $c=$c+1
    }
}
exit 0


$kamistarts   = New-Object System.Collections.Generic.List[float]
$kamiends     = New-Object System.Collections.Generic.List[float]
$shimostarts   = New-Object System.Collections.Generic.List[float]
$shimoends     = New-Object System.Collections.Generic.List[float]

$kamistarts.Add($starts[0]);$kamiends.Add($ends[0])
for ($i = 1; $i -lt $starts.Count; $i++) {
    $ma=$starts[$i] - $ends[$i-1]
    $duration=$ends[$i]-$starts[$i]
    if (($ma -ge 6 ) -and ($duration -le 13)) {
        $shimostarts.Add($starts[$i]);$shimoends.Add($ends[$i])
    }
    else{
        $kamistarts.Add($starts[$i]);$kamiends.Add($ends[$i])
    }
}
if(Test-Path segment.csv){Remove-Item segment.csv}
for($i = 0; $i -lt $kamistarts.Count; $i++){
    [string]$kamistarts[$i]+","+[string]$kamiends[$i]+",kami_"+($i+1) |Out-File -Append -Encoding utf8 segment.csv
}
for($i = 0; $i -lt $shimostarts.Count; $i++){
    [string]$shimostarts[$i]+","+[string]$shimoends[$i]+",shimo_"+($i+1) |Out-File -Append -Encoding utf8 segment.csv
}
exit 0
