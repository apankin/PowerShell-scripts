<#
.SYNOPSIS
    Сканирует папку (включая подпапки) и составляет отчёт по видеофайлам в CSV (и при желании в XLSX).

.DESCRIPTION
    Для каждого видеофайла определяет видеокодек (через ffprobe) и размер.
    Если указан целевой кодек, дополнительно считает ПРИБЛИЗИТЕЛЬНЫЙ размер после перекодирования
    по усреднённым коэффициентам эффективности кодеков. Точность невысокая (порядка +-30-50%),
    реальный размер зависит от содержимого видео и настроек качества.
    Оценка считается для всего файла целиком (звук в расчёте не выделяется).

.PARAMETER Path
    Папка для сканирования (подпапки обрабатываются автоматически).

.PARAMETER TargetCodec
    Необязательно. Кодек, в который планируется перекодировать:
    h264, h265 (hevc), av1, vp9, vp8, mpeg4 (xvid/divx), mpeg2, mpeg1, wmv, vc1, prores, dnxhd, mjpeg, theora.
    Если не указан, столбцы с оценкой в отчёт не добавляются.

.PARAMETER OutDir
    Папка для отчёта. По умолчанию рабочий стол.

.PARAMETER Excel
    Дополнительно создать файл XLSX (нужен модуль PowerShell ImportExcel).

.PARAMETER Delimiter
    Разделитель в CSV. По умолчанию ";" (подходит для русского Excel).

.EXAMPLE
    .\video_scan.ps1 -Path 'D:\Video'

.EXAMPLE
    .\video_scan.ps1 -Path 'D:\Video' -TargetCodec h265 -Excel
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string]$TargetCodec = '',
    [string]$OutDir = [Environment]::GetFolderPath('Desktop'),
    [switch]$Excel,
    [string]$Delimiter = ';'
)

# ---------- Проверки перед стартом ----------
if (-not (Get-Command ffprobe -ErrorAction SilentlyContinue)) {
    Write-Error 'ffprobe не найден. Установите FFmpeg (winget install Gyan.FFmpeg) и откройте PowerShell заново.'
    exit 1
}
if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Error "Папка не найдена: $Path"
    exit 1
}
if (-not (Test-Path -LiteralPath $OutDir -PathType Container)) {
    New-Item -ItemType Directory -Path $OutDir | Out-Null
}

# Проверяем модуль ImportExcel заранее, чтобы предупредить до долгого сканирования
$excel_ready = $false
if ($Excel) {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        $excel_ready = $true
    }
    else {
        Write-Warning 'Файл XLSX создать нельзя: не установлен модуль PowerShell ImportExcel. Установите его командой: Install-Module ImportExcel -Scope CurrentUser, затем запустите скрипт заново. CSV будет создан в любом случае.'
    }
}

# ---------- Справочники ----------
# Названия кодеков, которые можно указать в параметре TargetCodec -> имя кодека в ffprobe
$codec_aliases = @{
    'h264' = 'h264'; 'h.264' = 'h264'; 'avc' = 'h264'; 'x264' = 'h264'
    'h265' = 'hevc'; 'h.265' = 'hevc'; 'hevc' = 'hevc'; 'x265' = 'hevc'
    'av1' = 'av1'
    'vp9' = 'vp9'
    'vp8' = 'vp8'
    'mpeg4' = 'mpeg4'; 'xvid' = 'mpeg4'; 'divx' = 'mpeg4'
    'mpeg2' = 'mpeg2video'; 'mpeg2video' = 'mpeg2video'
    'mpeg1' = 'mpeg1video'; 'mpeg1video' = 'mpeg1video'
    'wmv' = 'wmv3'; 'wmv3' = 'wmv3'; 'wmv2' = 'wmv2'
    'vc1' = 'vc1'
    'theora' = 'theora'
    'mjpeg' = 'mjpeg'
    'prores' = 'prores'
    'dnxhd' = 'dnxhd'; 'dnxhr' = 'dnxhd'
}

# Относительный размер файла при одинаковом визуальном качестве (H.264 = 1.00).
# Значения усреднённые и приблизительные.
$size_factors = @{
    'h264' = 1.00
    'hevc' = 0.60
    'av1' = 0.45
    'vp9' = 0.65
    'vp8' = 1.20
    'mpeg4' = 1.50
    'msmpeg4v3' = 1.90
    'mpeg2video' = 2.20
    'mpeg1video' = 2.80
    'wmv3' = 1.60
    'wmv2' = 1.90
    'vc1' = 1.30
    'h263' = 2.00
    'flv1' = 1.80
    'theora' = 1.50
    'mjpeg' = 6.00
    'prores' = 10.00
    'dnxhd' = 9.00
}

$video_extensions = '.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.webm', '.m4v',
                    '.mpg', '.mpeg', '.ts', '.m2ts', '.mts', '.3gp', '.vob', '.ogv'

# Целевой кодек (если указан)
$target = ''
if ($TargetCodec) {
    $key = $TargetCodec.Trim().ToLower()
    if ($codec_aliases.ContainsKey($key)) {
        $target = $codec_aliases[$key]
    }
    else {
        Write-Error ("Неизвестный целевой кодек '{0}'. Доступные: {1}" -f $TargetCodec, (($codec_aliases.Keys | Sort-Object) -join ', '))
        exit 1
    }
}

$culture = [System.Globalization.CultureInfo]::CurrentCulture

# ---------- Сканирование ----------
$files = @(Get-ChildItem -LiteralPath $Path -Recurse -File |
    Where-Object { $video_extensions -contains $_.Extension.ToLower() })

if ($files.Count -eq 0) {
    Write-Host 'Видеофайлы в указанной папке не найдены.'
    exit 0
}

$rows = New-Object 'System.Collections.Generic.List[object]'
$i = 0

foreach ($file in $files) {
    $i++
    Write-Progress -Activity 'Анализ видео' -Status ("{0} из {1}: {2}" -f $i, $files.Count, $file.Name) -PercentComplete ($i / $files.Count * 100)

    # Читаем только заголовки файла: быстро, сам файл не изменяется
    $json_text = & ffprobe -v error -show_entries 'stream=codec_type,codec_name:stream_disposition=attached_pic' -of json -i $file.FullName 2>$null | Out-String
    $info = $null
    try { $info = $json_text | ConvertFrom-Json } catch { }

    # Берём первый видеопоток, обложки (attached_pic) пропускаем
    $video = $null
    if ($info -and $info.streams) {
        $video = @($info.streams | Where-Object { $_.codec_type -eq 'video' -and $_.disposition.attached_pic -ne 1 }) | Select-Object -First 1
    }

    $video_codec = ''
    $note = ''
    if ($video) { $video_codec = [string]$video.codec_name }
    else { $note = 'ffprobe не смог прочитать файл или в нём нет видео' }

    $size_mb = [math]::Round($file.Length / 1MB, 1)

    $row = [ordered]@{
        file_name   = $file.Name
        extension   = $file.Extension.TrimStart('.').ToLower()
        video_codec = $video_codec
        full_path   = $file.FullName
        size_mb     = [double]$size_mb
    }

    if ($target) {
        $est = $null
        $saving = $null

        if ($video_codec) {
            if ($video_codec -eq $target) {
                $est = [double]$size_mb
                $note = 'уже в целевом кодеке, перекодирование не нужно'
            }
            elseif (-not $size_factors.ContainsKey($video_codec)) {
                $note = "нет коэффициента для исходного кодека $video_codec, оценка невозможна"
            }
            else {
                $est = [math]::Round($file.Length / 1MB * $size_factors[$target] / $size_factors[$video_codec], 1)
                if ($est -ge $size_mb) { $note = 'перекодирование не уменьшит размер' }
            }
        }

        if ($null -ne $est -and $size_mb -gt 0) {
            $saving = [math]::Round((1 - $est / $size_mb) * 100, 1)
        }

        $row['target_codec'] = $target
        $row['estimated_size_mb'] = $est
        $row['estimated_saving_percent'] = $saving
    }

    $row['note'] = $note
    $rows.Add([pscustomobject]$row)
}

Write-Progress -Activity 'Анализ видео' -Completed

$rows = @($rows | Sort-Object size_mb -Descending)

# ---------- Вывод на экран ----------
if ($target) {
    $rows | Format-Table file_name, extension, video_codec, size_mb, estimated_size_mb, estimated_saving_percent, note -AutoSize
}
else {
    $rows | Format-Table file_name, extension, video_codec, size_mb, note -AutoSize
}

# Сводка по кодекам
$rows | Group-Object video_codec | ForEach-Object {
    [pscustomobject]@{
        video_codec = $_.Name
        files       = $_.Count
        size_gb     = [math]::Round(($_.Group | Measure-Object size_mb -Sum).Sum / 1024, 2)
    }
} | Sort-Object size_gb -Descending | Format-Table -AutoSize

# ---------- CSV ----------
# Числа превращаем в текст по правилам текущего языка Windows (запятая в русской локали),
# чтобы Excel распознал их как числа
$csv_rows = foreach ($r in $rows) {
    $copy = [ordered]@{}
    foreach ($p in $r.PSObject.Properties) {
        if ($p.Value -is [double]) { $copy[$p.Name] = $p.Value.ToString('0.0', $culture) }
        else { $copy[$p.Name] = $p.Value }
    }
    [pscustomobject]$copy
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csv_path = Join-Path $OutDir "video_report_$stamp.csv"
$csv_rows | Export-Csv -LiteralPath $csv_path -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8
Write-Host "CSV сохранён: $csv_path"

# ---------- XLSX (по желанию) ----------
if ($excel_ready) {
    $xlsx_path = Join-Path $OutDir "video_report_$stamp.xlsx"
    Import-Module ImportExcel
    $rows | Export-Excel -Path $xlsx_path -WorksheetName 'video' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    Write-Host "XLSX сохранён: $xlsx_path"
}
