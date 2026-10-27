<#
.SYNOPSIS
    Перекодирует видеофайлы в указанный кодек и пишет текстовый отчёт (результат, размер до и после).

.DESCRIPTION
    Вход: либо путь к одному видеофайлу, либо текстовый файл (.txt) со списком полных путей
    (один путь на строку, кодировка UTF-8, пустые строки и строки с # в начале игнорируются).
    Звук, субтитры, главы и метаданные копируются без изменений.

    Режимы сохранения:
      по умолчанию  - новый файл рядом с оригиналом с суффиксом (video_h265.mkv), оригинал не трогается;
      -ReplaceOriginal - после успешной проверки оригинал отправляется в Корзину,
                         а новый файл получает его имя. Если результат не меньше оригинала, замены не будет.

.PARAMETER InputPath
    Полный путь к видеофайлу или к .txt со списком файлов.

.PARAMETER Codec
    Целевой кодек: h264, h265 (hevc), av1, vp9.

.PARAMETER Container
    Необязательно. Контейнер результата: mkv, mp4, m4v, mov, webm.
    Если не указан, остаётся исходный. Если исходный контейнер не подходит (avi, wmv, flv и т.п.)
    или в нём не поместятся дорожки, скрипт сам выберет mkv и напишет об этом в отчёте.

.PARAMETER Encoder
    gpu (по умолчанию, видеокарта NVIDIA, NVENC) или cpu (процессор, лучше сжимает, но намного медленнее).

.PARAMETER Crf
    Необязательно. Число качества: меньше = лучше качество и больше файл.
    Значения по умолчанию: cpu: h264=20, h265=22, av1=30, vp9=32; gpu: h264=21, h265=24, av1=28.

.PARAMETER Preset
    Необязательно. Пресет скорости.
    cpu x264/x265: ultrafast ... medium ... slow ... veryslow (по умолчанию medium); cpu av1: 0-13 (по умолчанию 6);
    gpu: p1 (быстро) ... p7 (лучшее качество, по умолчанию).

.PARAMETER ReplaceOriginal
    Заменять оригинал (старый файл уходит в Корзину).

.PARAMETER ReportPath
    Путь к текстовому отчёту. По умолчанию рабочий стол.

.EXAMPLE
    .\video_transcode.ps1 -InputPath 'D:\Video\film.mp4' -Codec h265

.EXAMPLE
    .\video_transcode.ps1 -InputPath 'D:\list.txt' -Codec h265 -Container mkv -Encoder cpu -Preset slow
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [Parameter(Mandatory = $true)]
    [string]$Codec,
    [string]$Container = '',
    [ValidateSet('cpu', 'gpu')]
    [string]$Encoder = 'gpu',
    [int]$Crf = -1,
    [string]$Preset = '',
    [switch]$ReplaceOriginal,
    [string]$ReportPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) ("transcode_report_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
)

# Чтобы русские имена файлов в сообщениях ffmpeg не превращались в кракозябры
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$culture = [System.Globalization.CultureInfo]::CurrentCulture

# ---------- Проверки перед стартом ----------
foreach ($tool in 'ffmpeg', 'ffprobe') {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
        Write-Error "$tool не найден. Установите FFmpeg (winget install Gyan.FFmpeg) и откройте PowerShell заново."
        exit 1
    }
}

$codec_aliases = @{
    'h264' = 'h264'; 'h.264' = 'h264'; 'avc' = 'h264'; 'x264' = 'h264'
    'h265' = 'hevc'; 'h.265' = 'hevc'; 'hevc' = 'hevc'; 'x265' = 'hevc'
    'av1' = 'av1'
    'vp9' = 'vp9'
}
$codec_key = $Codec.Trim().ToLower()
if (-not $codec_aliases.ContainsKey($codec_key)) {
    Write-Error "Неизвестный кодек '$Codec'. Доступные: h264, h265 (hevc), av1, vp9."
    exit 1
}
$target_codec = $codec_aliases[$codec_key]

$suffix_names = @{ 'h264' = 'h264'; 'hevc' = 'h265'; 'av1' = 'av1'; 'vp9' = 'vp9' }
$suffix = $suffix_names[$target_codec]

if ($Encoder -eq 'gpu' -and $target_codec -eq 'vp9') {
    Write-Error 'Видеокарты NVIDIA не умеют кодировать VP9. Используйте -Encoder cpu или другой кодек.'
    exit 1
}

# Значения качества и пресета по умолчанию: @(crf, preset)
$defaults = @{
    'cpu' = @{ 'h264' = @(20, 'medium'); 'hevc' = @(22, 'medium'); 'av1' = @(30, '6'); 'vp9' = @(32, '') }
    'gpu' = @{ 'h264' = @(21, 'p7'); 'hevc' = @(24, 'p7'); 'av1' = @(28, 'p7') }
}
$crf_value = $defaults[$Encoder][$target_codec][0]
$preset_value = $defaults[$Encoder][$target_codec][1]
if ($Crf -ge 0) { $crf_value = $Crf }
if ($Preset) { $preset_value = $Preset }

# Параметры видеокодера
$enc_args = @()
if ($Encoder -eq 'cpu') {
    switch ($target_codec) {
        'hevc' { $enc_args = @('-c:v', 'libx265', '-crf', [string]$crf_value, '-preset', $preset_value, '-x265-params', 'log-level=error') }
        'h264' { $enc_args = @('-c:v', 'libx264', '-crf', [string]$crf_value, '-preset', $preset_value, '-pix_fmt', 'yuv420p') }
        'av1' { $enc_args = @('-c:v', 'libsvtav1', '-crf', [string]$crf_value, '-preset', $preset_value) }
        'vp9' { $enc_args = @('-c:v', 'libvpx-vp9', '-crf', [string]$crf_value, '-b:v', '0', '-row-mt', '1') }
    }
}
else {
    $nvenc = @('-preset', $preset_value, '-tune', 'hq', '-rc', 'vbr', '-cq', [string]$crf_value, '-b:v', '0', '-multipass', 'fullres')
    switch ($target_codec) {
        'hevc' { $enc_args = @('-c:v', 'hevc_nvenc') + $nvenc }
        'h264' { $enc_args = @('-c:v', 'h264_nvenc') + $nvenc + @('-profile:v', 'high', '-pix_fmt', 'yuv420p') }
        'av1' { $enc_args = @('-c:v', 'av1_nvenc') + $nvenc }
    }
}

# ---------- Список файлов ----------
if (Test-Path -LiteralPath $InputPath -PathType Container) {
    Write-Error 'Указана папка. Укажите файл или .txt со списком файлов (список можно получить скриптом video_scan.ps1).'
    exit 1
}
if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    Write-Error "Файл не найден: $InputPath"
    exit 1
}
$input_item = Get-Item -LiteralPath $InputPath
if ($input_item.Extension -ieq '.txt') {
    $files = @(Get-Content -LiteralPath $InputPath -Encoding UTF8 |
        ForEach-Object { $_.Trim().Trim('"') } |
        Where-Object { $_ -and -not $_.StartsWith('#') })
}
else {
    $files = @($input_item.FullName)
}
if ($files.Count -eq 0) {
    Write-Host 'В списке нет ни одного пути.'
    exit 0
}

# ---------- Вспомогательные функции ----------
function Write-Report {
    param([string]$Line)
    Add-Content -LiteralPath $ReportPath -Value $Line -Encoding UTF8
}

function Format-Mb {
    param([double]$Bytes)
    return ($Bytes / 1MB).ToString('0.0', $culture) + ' МБ'
}

function New-Result {
    param([string]$Status, [string]$Message, [double]$Old, [double]$New)
    return [pscustomobject]@{ status = $Status; message = $Message; old = $Old; new = $New }
}

# Читает из файла: видеокодек, список аудио- и субтитровых кодеков, длительность
function Get-MediaInfo {
    param([string]$FilePath)
    $json_text = & ffprobe -v error -show_entries 'format=duration:stream=codec_type,codec_name:stream_disposition=attached_pic' -of json -i $FilePath 2>$null | Out-String
    $json = $null
    try { $json = $json_text | ConvertFrom-Json } catch { return $null }
    if (-not $json -or -not $json.streams) { return $null }

    $video = @($json.streams | Where-Object { $_.codec_type -eq 'video' -and $_.disposition.attached_pic -ne 1 }) | Select-Object -First 1
    $duration = 0.0
    if ($json.format -and $json.format.duration) { $duration = [double]$json.format.duration }

    return [pscustomobject]@{
        video_codec  = if ($video) { [string]$video.codec_name } else { '' }
        audio_codecs = @($json.streams | Where-Object { $_.codec_type -eq 'audio' } | ForEach-Object { [string]$_.codec_name })
        sub_codecs   = @($json.streams | Where-Object { $_.codec_type -eq 'subtitle' } | ForEach-Object { [string]$_.codec_name })
        duration     = $duration
    }
}

# Выбирает контейнер: берёт нужный, а если в него не поместятся дорожки, переключается на mkv
function Resolve-Container {
    param([string]$Requested, [string]$SourceExt, $Info)

    $allowed = @('mkv', 'mp4', 'm4v', 'mov', 'webm')
    if ($Requested) { $container = $Requested.TrimStart('.').ToLower() }
    else { $container = $SourceExt.TrimStart('.').ToLower() }

    $reason = ''
    if ($allowed -notcontains $container) {
        $reason = "контейнер .$container не подходит для выбранного кодека"
    }
    elseif (@('mp4', 'm4v', 'mov') -contains $container) {
        $bad_audio = @($Info.audio_codecs | Where-Object { @('truehd', 'mlp', 'dts') -contains $_ -or $_ -like 'pcm_*' })
        $bad_subs = @($Info.sub_codecs | Where-Object { $_ -ne 'mov_text' })
        if ($bad_audio.Count -gt 0) { $reason = "звук $($bad_audio -join ', ') не помещается в $container" }
        elseif ($bad_subs.Count -gt 0) { $reason = "субтитры $($bad_subs -join ', ') не помещаются в $container" }
    }
    elseif ($container -eq 'webm') {
        $bad_audio = @($Info.audio_codecs | Where-Object { @('opus', 'vorbis') -notcontains $_ })
        if (@('vp9', 'av1') -notcontains $target_codec) { $reason = 'webm поддерживает только VP9 и AV1' }
        elseif ($bad_audio.Count -gt 0) { $reason = "звук $($bad_audio -join ', ') не помещается в webm" }
        elseif ($Info.sub_codecs.Count -gt 0) { $reason = 'субтитры не помещаются в webm' }
    }

    $note = ''
    if ($reason) {
        $note = "контейнер заменён на mkv: $reason"
        $container = 'mkv'
    }
    return [pscustomobject]@{ container = $container; note = $note }
}

# Подсказка по типичным ошибкам ffmpeg
function Get-ErrorHint {
    param([string]$Text)
    if ($Text -match 'nvenc|No capable devices|Cannot load|OpenEncodeSessionEx|Driver does not support') {
        return ' Подсказка: проверьте драйвер NVIDIA; кодирование AV1 на видеокарте работает только на RTX 40xx и новее.'
    }
    if ($Text -match 'codec not currently supported in container|Subtitle|Could not write header') {
        return ' Подсказка: попробуйте указать -Container mkv.'
    }
    return ''
}

# Кодирует один файл и возвращает результат
function Invoke-Transcode {
    param([string]$SourcePath)

    $item = Get-Item -LiteralPath $SourcePath -ErrorAction SilentlyContinue
    if (-not $item -or $item.PSIsContainer) { return New-Result 'ERROR' 'файл не найден' 0 0 }
    $old_size = [double]$item.Length

    $info = Get-MediaInfo -FilePath $item.FullName
    if (-not $info -or -not $info.video_codec) {
        return New-Result 'ERROR' 'ffprobe не смог прочитать файл или в нём нет видео' $old_size 0
    }
    if ($info.video_codec -eq $target_codec) {
        return New-Result 'SKIP' "уже в кодеке $($info.video_codec), перекодирование не нужно" $old_size 0
    }

    $resolved = Resolve-Container -Requested $Container -SourceExt $item.Extension -Info $info
    $ext = '.' + $resolved.container
    $dir = $item.DirectoryName
    $base = [System.IO.Path]::GetFileNameWithoutExtension($item.Name)
    $temp_path = Join-Path $dir ("{0}.tmp_{1}{2}" -f $base, $suffix, $ext)

    if ($ReplaceOriginal) {
        $final_path = Join-Path $dir ($base + $ext)
        if ((Test-Path -LiteralPath $final_path) -and ($final_path -ine $item.FullName)) {
            return New-Result 'ERROR' "файл с таким именем уже существует: $final_path" $old_size 0
        }
    }
    else {
        $final_path = Join-Path $dir ("{0}_{1}{2}" -f $base, $suffix, $ext)
        if (Test-Path -LiteralPath $final_path) {
            return New-Result 'SKIP' 'выходной файл уже существует' $old_size 0
        }
    }

    # Остаток от прерванного запуска удаляем
    if (Test-Path -LiteralPath $temp_path) { Remove-Item -LiteralPath $temp_path -Force }

    # Собираем команду ffmpeg
    $ff_args = @('-hide_banner', '-loglevel', 'error', '-nostdin', '-y', '-i', $item.FullName,
        '-map', '0:V', '-map', '0:a?', '-map', '0:s?')
    if ($resolved.container -eq 'mkv') { $ff_args += @('-map', '0:t?') }
    $ff_args += @('-map_metadata', '0', '-map_chapters', '0')
    $ff_args += $enc_args
    $ff_args += @('-c:a', 'copy', '-c:s', 'copy')
    if ($target_codec -eq 'hevc' -and @('mp4', 'm4v', 'mov') -contains $resolved.container) {
        $ff_args += @('-tag:v', 'hvc1')
    }
    $ff_args += $temp_path

    $ff_out = & ffmpeg @ff_args 2>&1 | Out-String
    $exit_code = $LASTEXITCODE

    if ($exit_code -ne 0 -or -not (Test-Path -LiteralPath $temp_path)) {
        if (Test-Path -LiteralPath $temp_path) { Remove-Item -LiteralPath $temp_path -Force }
        $tail = (($ff_out -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 2) -join ' / ')
        return New-Result 'ERROR' ("ffmpeg завершился с ошибкой: $tail" + (Get-ErrorHint $ff_out)) $old_size 0
    }

    # Проверяем результат: кодек и длительность должны совпасть с ожидаемыми
    $new_size = [double](Get-Item -LiteralPath $temp_path).Length
    $info_new = Get-MediaInfo -FilePath $temp_path
    $problem = ''
    if ($new_size -le 0) { $problem = 'выходной файл пустой' }
    elseif (-not $info_new -or $info_new.video_codec -ne $target_codec) { $problem = 'в выходном файле не тот кодек или он не читается' }
    elseif ($info.duration -gt 0 -and [math]::Abs($info_new.duration - $info.duration) -gt [math]::Max(2, $info.duration * 0.01)) {
        $problem = 'длительность результата не совпадает с оригиналом'
    }
    if ($problem) {
        Remove-Item -LiteralPath $temp_path -Force
        return New-Result 'ERROR' "проверка результата не пройдена: $problem" $old_size 0
    }

    $note = $resolved.note
    $is_larger = $new_size -ge $old_size

    if ($ReplaceOriginal) {
        if ($is_larger) {
            Remove-Item -LiteralPath $temp_path -Force
            return New-Result 'LARGER' 'результат не меньше оригинала, оригинал оставлен без изменений, результат удалён' $old_size $new_size
        }
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($item.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin')
            Move-Item -LiteralPath $temp_path -Destination $final_path
        }
        catch {
            return New-Result 'ERROR' "не удалось заменить оригинал: $($_.Exception.Message). Готовый файл остался: $temp_path" $old_size $new_size
        }
        (Get-Item -LiteralPath $final_path).LastWriteTime = $item.LastWriteTime
        $msg = "оригинал отправлен в Корзину, новый файл: $final_path"
        if ($note) { $msg += "; $note" }
        return New-Result 'OK' $msg $old_size $new_size
    }

    Move-Item -LiteralPath $temp_path -Destination $final_path
    (Get-Item -LiteralPath $final_path).LastWriteTime = $item.LastWriteTime
    $msg = "новый файл: $final_path"
    if ($note) { $msg += "; $note" }
    if ($is_larger) {
        return New-Result 'LARGER' "результат БОЛЬШЕ оригинала (файл сохранён, решите сами, нужен ли он); $msg" $old_size $new_size
    }
    return New-Result 'OK' $msg $old_size $new_size
}

# ---------- Заголовок отчёта ----------
$mode_text = 'новые файлы рядом с оригиналами (суффикс _' + $suffix + '), оригиналы не трогаются'
if ($ReplaceOriginal) { $mode_text = 'замена оригиналов (старый файл уходит в Корзину)' }
$container_text = 'как у исходного файла'
if ($Container) { $container_text = $Container }

Write-Report ('Отчёт о перекодировании, ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Write-Report ("Кодек: $target_codec | кодировщик: $Encoder | качество (crf/cq): $crf_value | пресет: $preset_value | контейнер: $container_text")
Write-Report "Режим: $mode_text"
Write-Report ("Файлов к обработке: " + $files.Count)
Write-Report ('-' * 100)

# ---------- Основной цикл ----------
$counters = @{ OK = 0; LARGER = 0; SKIP = 0; ERROR = 0 }
$total_old = 0.0
$total_new = 0.0
$colors = @{ OK = 'Green'; LARGER = 'Yellow'; SKIP = 'Gray'; ERROR = 'Red' }
$total_timer = [System.Diagnostics.Stopwatch]::StartNew()
$i = 0

foreach ($src in $files) {
    $i++
    Write-Progress -Activity 'Перекодирование' -Status ("{0} из {1}: {2}" -f $i, $files.Count, (Split-Path -Leaf $src)) -PercentComplete ($i / $files.Count * 100)
    Write-Host ("[{0}/{1}] {2}" -f $i, $files.Count, $src)

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $r = Invoke-Transcode -SourcePath $src
    }
    catch {
        $r = New-Result 'ERROR' ("непредвиденная ошибка: " + $_.Exception.Message) 0 0
    }
    $timer.Stop()
    $counters[$r.status]++

    if ($r.status -eq 'OK') {
        $total_old += $r.old
        $total_new += $r.new
    }

    $line = "[{0}] {1}" -f $r.status, $src
    if ($r.old -gt 0) { $line += ' | было: ' + (Format-Mb $r.old) }
    if ($r.new -gt 0) {
        $line += ' | стало: ' + (Format-Mb $r.new)
        $line += ' | изменение: ' + ((($r.new / $r.old) - 1) * 100).ToString('+0.0;-0.0', $culture) + '%'
    }
    if ($r.status -eq 'OK' -or $r.status -eq 'LARGER') {
        $line += ' | время: ' + $timer.Elapsed.TotalMinutes.ToString('0.0', $culture) + ' мин'
    }
    $line += ' | ' + $r.message
    Write-Report $line
    Write-Host ("    " + $line) -ForegroundColor $colors[$r.status]
}

Write-Progress -Activity 'Перекодирование' -Completed
$total_timer.Stop()

# ---------- Итоги ----------
$saved_gb = ($total_old - $total_new) / 1GB
Write-Report ('-' * 100)
Write-Report ("Итого: успешно (OK) = {0}, результат больше оригинала (LARGER) = {1}, пропущено (SKIP) = {2}, ошибок (ERROR) = {3}" -f $counters.OK, $counters.LARGER, $counters.SKIP, $counters.ERROR)
Write-Report ("По успешным файлам: было {0} ГБ, стало {1} ГБ, сэкономлено {2} ГБ" -f
    ($total_old / 1GB).ToString('0.00', $culture), ($total_new / 1GB).ToString('0.00', $culture), $saved_gb.ToString('0.00', $culture))
Write-Report ("Общее время: " + $total_timer.Elapsed.TotalMinutes.ToString('0.0', $culture) + ' мин')

Write-Host ''
Write-Host "Готово. Отчёт: $ReportPath" -ForegroundColor Cyan
