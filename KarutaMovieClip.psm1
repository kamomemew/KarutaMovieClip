function Get-Percentile {
# MIT License

# Copyright (c) 2022 Jim Birley

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
<#
.SYNOPSIS
    Returns the specified percentile value for a given set of numbers.
 
.DESCRIPTION
    This function expects a set of numbers passed as an array to the 'Sequence' parameter.  For a given percentile, passed as the 'Percentile' argument,
    it returns the calculated percentile value for the set of numbers.
 
.PARAMETER Sequence
    A array of integer and/or decimal values the function uses as the data set.
.PARAMETER Percentile
    The target percentile to be used by the function's algorithm. 
 
.EXAMPLE
    $values = 98.2,96.5,92.0,97.8,100,95.6,93.3
    Get-Percentile -Sequence $values -Percentile 0.95
 
.NOTES
    Author:  Jim Birley
#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] 
        [Double[]]$Sequence
        ,
        [Parameter(Mandatory)]
        [Double]$Percentile
    )
   
    $Sequence = $Sequence | Sort-Object
    [int]$N = $Sequence.Length
    Write-Verbose "N is $N"
    [Double]$Num = ($N - 1) * $Percentile + 1
    Write-Verbose "Num is $Num"
    if ($num -eq 1) {
        return $Sequence[0]
    } elseif ($num -eq $N) {
        return $Sequence[$N-1]
    } else {
        $k = [Math]::Floor($Num)
        Write-Verbose "k is $k"
        [Double]$d = $num - $k
        Write-Verbose "d is $d"
        return $Sequence[$k - 1] + $d * ($Sequence[$k] - $Sequence[$k - 1])
    }
}
function Merge-Gaps {
    param (
        [System.Collections.Generic.List[float]]$starts,
        [System.Collections.Generic.List[float]]$ends,
        [float]$sec
    )
    process {
        $processed_starts = New-Object System.Collections.Generic.List[float]
        $processed_ends = New-Object System.Collections.Generic.List[float]

        $processed_starts.Add($starts[0])
        $starts = $starts[1..($starts.Count - 1)]
        for ($i = 0; $i -lt $starts.Count; $i++) {
            if ( ($starts[$i] - $ends[$i]) -gt $sec) {
                $processed_starts.Add($starts[$i])
                $processed_ends.Add($ends[$i])
            }
        }
        $processed_ends.Add($ends[-1])
        return $processed_starts, $processed_ends
    }
}

function Select-Span {
    param (
        [System.Collections.Generic.List[float]]$starts,
        [System.Collections.Generic.List[float]]$ends,
        [float]$lt=0,
        [float]$gt=0
    )
    process {
        $processed_starts=New-Object System.Collections.Generic.List[float]
        $processed_ends=New-Object System.Collections.Generic.List[float]
        if ($gt) {
            for ($i = 0; $i -lt $starts.Count; $i++) {
                if (($ends[$i] - $starts[$i]) -ge  $gt) {
                    $processed_starts.Add($starts[$i])
                    $processed_ends.Add($ends[$i])
                }
            }
            return $processed_starts,$processed_ends
        }
        elseif ($lt) {
            for ($i = 0; $i -lt $starts.Count; $i++) {
                if (($ends[$i] - $starts[$i]) -lt  $lt) {
                    $processed_starts.Add($starts[$i])
                    $processed_ends.Add($ends[$i])
                }
            }
            return $processed_starts,$processed_ends
        }
        else{
            return $starts,$ends
        }
    }
}

function Deny-Gap {
    param(
        [switch]$before=$false,
        [switch]$after=$false,
        [System.Collections.Generic.List[float]]$starts,
        [System.Collections.Generic.List[float]]$ends,
        [float]$sec
    )
    process {
        if($before -and $after){
            $processed_starts=New-Object System.Collections.Generic.List[float]
            $processed_ends=New-Object System.Collections.Generic.List[float]
            # index[0]
            if(($starts[0] -le $sec) -or (($starts[1] - $ends[0]) -le $sec)){
                $processed_starts.Add($starts[0]);$processed_ends.Add($ends[0])
            }
            # index[0...N-1]
            for ($i = 1; $i -lt $starts.Count-1; $i++) {
                if((($starts[$i]-$ends[$i-1]) -le $sec) -or (($starts[$i+1] - $ends[$i]) -le $sec)){
                    $processed_starts.Add($starts[$i]);$processed_ends.Add($ends[$i])
                }
            }
            # index[N]
            if(($starts[$i]-$ends[$i-1]) -le $sec){
                $processed_starts.Add($starts[$i]);$processed_ends.Add($ends[$i])
            }
            return $processed_starts,$processed_ends
        }
        elseif($before){
            $processed_starts=New-Object System.Collections.Generic.List[float]
            $processed_ends=New-Object System.Collections.Generic.List[float]
            # index[0] ??
            $processed_starts.Add($starts[0]);$processed_ends.Add($ends[0])
            # index[1...N]
            for ($i = 1; $i -lt $starts.Count; $i++) {
                if(($starts[$i]-$ends[$i-1]) -le $sec){
                    $processed_starts.Add($starts[$i]);$processed_ends.Add($ends[$i])
                }
            }
            return $processed_starts,$processed_ends
        }
        elseif ($after) {
            <# Action when this condition is true #>
        }
        else{
            return $starts,$ends
        }
    }
}

function Select-Invert {
    param (
        [System.Collections.Generic.List[float]]$starts,
        [System.Collections.Generic.List[float]]$ends
    )
    process {
        $processed_starts=New-Object System.Collections.Generic.List[float]
        $processed_ends=New-Object System.Collections.Generic.List[float]
        $processed_starts.Add(0)
        for ($i = 0; $i -lt ($starts.Count-1); $i++) {
            $processed_ends.Add($starts[$i])
            $processed_starts.Add($ends[$i])
        }
        $processed_ends.Add($starts[-1])
        if (($processed_starts[0] -eq 0) -and ($processed_ends[0] -eq 0)) {
            $processed_ends=$processed_ends[1..($processed_ends.Count-1)]
            $processed_starts=$processed_starts[1..($processed_starts.Count-1)]
        }
        return $processed_starts,$processed_ends
    }
}

function Test-DataIntegrity {
    param (
        [System.Collections.Generic.List[float]]$starts,
        [System.Collections.Generic.List[float]]$ends
    )
    process {
        if ($starts.Count -eq $ends.Count){
            for ($i = 0; $i -lt $starts.Count; $i++) {
                if ( -not ($ends[$i] -ge $starts[$i])) {
                    return $false
                }
                if ($i -gt 0){
                    if ( -not (($starts[$i]-$starts[$i-1] -gt 0) -and ($ends[$i]-$ends[$i-1] -gt 0))){
                        return $false
                    }
                }
            }
            return $true
        }
        else {
            return $false
        }
    }
}

function Get-KamiShimoList {
        param (
        [System.Collections.Generic.List[float]]$starts,
        [System.Collections.Generic.List[float]]$ends
    )
    process{
        $kamishimo = New-Object  "System.Collections.Generic.List[string]"
        for ($i = 0; $i -lt $starts.Count; $i++) {
            $kamishimo.Add("")
        }
        $kamishimo[0]="?";$kamishimo[-1]=("?")
        for ($i = 1; $i -lt $starts.Count-1; $i++) {
            if((($ends[$i-1]-$starts[$i]) -ge 6) -and (($starts[$i+1] - $ends[$i]) -ge 6) -and (($ends[$i] - $starts[$i]) -le 13)){
                $kamishimo[$i] = "?"
                continue
            }
            if (($starts[$i+1] - $ends[$i]) -ge 6){
                $kamishimo[$i] = "kami"
                continue
            }
            elseif(($starts[$i]-$ends[$i-1]) -ge 6) {
                $kamishimo[$i] = "shimo"
            }
        }
        $modified=$false
        do {
            $modified = $false
            for ($i = 1; $i -lt $starts.Count-1; $i++) {
                if($kamishimo[$i] -eq ""){
                    if ($kamishimo[$i+1] |Select-String "kami" -Quiet){
                        $kamishimo[$i] = "*shimo"
                        $modified = $true
                    }
                    elseif ($kamishimo[$i-1] |Select-String "shimo" -Quiet) {
                        $kamishimo[$i] = "*kami"
                        $modified = $true
                    }
                }
            }
        } while ($modified)
        for ($i = 1; $i -lt $starts.Count-1; $i++) {
            if($kamishimo[$i] -eq ""){
                $kamishimo[$i] = "?"
            }
        }
        return $kamishimo
    }
}

function Test-Silence {
    <#
    .SYNOPSIS
        無音かどうかを判定する
    .PARAMETER starts
        「有音部の」始まりの秒数のリスト
    .PARAMETER ends
        「有音部の」終わりの秒数のリスト
    #>
    param (
        [System.Collections.Generic.List[float]]$starts,
        [System.Collections.Generic.List[float]]$ends,
        [float]$sec
    )
    process{
        for ($i = 0; $i -lt $starts.Count; $i++) {
            if ($starts[$i] -gt $sec){
                return $true
            }
            elseif ($ends[$i] -lt $sec) {
                continue
            }
            elseif (($starts[$i] -le $sec) -and ($ends[$i] -ge $sec)) {
                return $false
            }
        }
        return $true
    }
}

function Get-SlopeFactor {
    <#
    .SYNOPSIS
        音量変化と包絡線の二乗差を求める.
        戻り値は時間で規格化される
    .PARAMETER times
        時間のリスト
    .PARAMETER volumes
        ボリュームのリスト(時間と同じ長さであること)
    .PARAMETER xy1
        区間開始時の@(時間,音量)
    .PARAMETER xy2
        区間終了時の@(時間,音量)
    #>
    param (
        [System.Collections.Generic.List[float]]$times,
        [System.Collections.Generic.List[float]]$volumes,
        [System.Collections.Generic.List[float]]$start,
        [System.Collections.Generic.List[float]]$end
    )
    process{
        $x1=$start[0]; $y1=$start[1]
        $x2=$end[0]; $y2=$end[1]
        $linior={param($x); ($y2-$y1)/($x2-$x1)*($x-$x1) + $y1}
        $deltas=New-Object System.Collections.Generic.List[float]
        for ($i = 0; $i -lt $times.Count; $i++) {
            if(($times[$i] -ge $x1) -and ($times[$i] -le $x2)){
                $d=[Math]::Pow((&$linior $times[$i])-$volumes[$i],2)
                $deltas.Add($d)
            }
            elseif ($times[$i] -gt $x2) {
                break
            }
        }
        $deltaInfo=$deltas|Measure-Object -Sum
        return ($deltaInfo.Sum/$deltaInfo.Count)/($x2-$x1)
    }
}
function Get-VolumeInfo {
    param (
        [System.Collections.Generic.List[float]]$times,
        [System.Collections.Generic.List[float]]$volumes,
        [float]$start,
        [float]$end
    )
    process{
        if (-not ($times.Count -eq $volumes.Count)) {
            return -1
        }
        $period=New-Object System.Collections.Generic.List[float]
        for ($i = 0; $i -lt $times.Count; $i++) {
            if (($times[$i] -ge $start) -and ($times[$i] -lt $end)){
                $period.Add($volumes[$i])
            }
            elseif ($times[$i] -ge $end ) {
                break
            }
        }
        $InfoHash=@{}
        $VolumeInfo=$period|Measure-Object -Maximum -Minimum -Average
        $InfoHash.Add("DynamicRange",$VolumeInfo.Maximum-$VolumeInfo.Minimum)
        $InfoHash.Add("Maximum",$VolumeInfo.Maximum)
        $InfoHash.Add("Minimum",$VolumeInfo.Minimum)
        $InfoHash.Add("Average",$VolumeInfo.Average)
        return $InfoHash
    }
}

Export-ModuleMember -Function Merge-Gaps
Export-ModuleMember -Function Select-Span
Export-ModuleMember -Function Deny-Gap
Export-ModuleMember -Function Select-Invert
Export-ModuleMember -Function Test-DataIntegrity
Export-ModuleMember -Function Test-Silence
Export-ModuleMember -Function Get-SlopeFactor
Export-ModuleMember -Function Get-VolumeInfo
Export-ModuleMember -Function Get-Percentile
Export-ModuleMember -Function Get-KamiShimoList