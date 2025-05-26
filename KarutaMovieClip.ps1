Import-Module ".\KarutaMovieClip.psm1"
$configure=Get-Content -Path "configure.json" -Encoding "utf8" | ConvertFrom-Json
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
$volume_raw_text=&$configure.ffmpeg -v error -ss 600  -t 360 -i $movie -af "highpass=f=200,lowpass=f=3000,afftdn,aresample=44100,asetnsamples=2205,astats=reset=1:metadata=1,ametadata=print:key=lavfi.astats.Overall.peak_level:file='pipe\:1'" -vn -f null - |& { process{ $_.ToString() }}
$peaks=$volume_raw_text| select-String "peak_level=(.*)$" | &{ process {[float]$($_.matches.groups[1]).ToString() }}
$silence_threshold=(Get-Percentile -Sequence $peaks -Percentile 0.7).ToString("0.00")
# FFmpegで空白検出、時間がかかる
$silence_raw_text=&$configure.ffmpeg -i $movie -af "highpass=f=200,lowpass=f=3000,afftdn,silencedetect=n=$($silence_threshold)dB:d=1.5:m=0" -vn -f null - 2>&1|ForEach-Object { $_.ToString() }
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
