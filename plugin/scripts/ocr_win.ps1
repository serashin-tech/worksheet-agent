# Windows 내장 OCR(Windows.Media.Ocr, 한국어)로 이미지의 글자 줄과 위치를 읽는다.
# 사용: powershell.exe -NoProfile -File ocr_win.ps1 <list.txt> <out.json>
#   list.txt: 한 줄에 이미지 경로 하나(UTF-8). out.json: [{"path":..., "lines":[{"text":..., "y":..., "h":...}]}]
# Windows PowerShell 5.1에서만 돈다(PowerShell 7에는 WinRT 형식이 없음).
param([string]$List, [string]$Out)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Wait-Op($op, [type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(); $task.Result }
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new("ko"))
if (-not $engine) { throw "Korean OCR language is not installed" }
$res = @()
foreach ($p in [IO.File]::ReadAllLines($List, [Text.Encoding]::UTF8)) {
  if (-not $p) { continue }
  $f = Wait-Op ([Windows.Storage.StorageFile]::GetFileFromPathAsync($p)) ([Windows.Storage.StorageFile])
  $s = Wait-Op ($f.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  $d = Wait-Op ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($s)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $b = Wait-Op ($d.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $r = Wait-Op ($engine.RecognizeAsync($b)) ([Windows.Media.Ocr.OcrResult])
  $lines = @($r.Lines | ForEach-Object { $w = @($_.Words); [ordered]@{ text = $_.Text; x = [int]$w[0].BoundingRect.X; y = [int]$w[0].BoundingRect.Y; h = [int]$w[0].BoundingRect.Height } })
  $res += [ordered]@{ path = $p; lines = $lines }
  $s.Dispose()
}
[IO.File]::WriteAllText($Out, (ConvertTo-Json -InputObject $res -Depth 5 -Compress), (New-Object Text.UTF8Encoding $false))
