#Requires -Version 7.0
<#
.SYNOPSIS
    Deletes folders (prefixes) from an AWS S3 bucket after showing exactly what will be deleted.

.DESCRIPTION
    1. Lists the folders in the bucket (or inside -Prefix).
    2. Asks which folders to delete (numbers: 1,3,5-7 or all).
    3. Lists every object inside each selected folder, with sizes and totals.
    4. Asks for confirmation (you have to type DELETE).
    5. Deletes exactly the objects that were shown and checks what is left in the folders.

    Objects uploaded to a folder after it was listed are NOT deleted: the script reports them instead.

.EXAMPLE
    .\Remove-S3Folders.ps1
    Interactive run on the default bucket from $Defaults.

.EXAMPLE
    .\Remove-S3Folders.ps1 -Bucket my-bucket -Prefix archive/2025 -DryRun
    Show the folders inside archive/2025/ and what would be deleted, without deleting anything.
#>
[CmdletBinding()]
param(
    [string]$Bucket,
    [string]$Prefix = '',      # родительская «папка», внутри которой показывать папки; пусто = корень бакета
    [string]$AwsProfile,
    [string]$EndpointUrl,
    [switch]$DryRun
)

# =====================================================================================
#  ЗНАЧЕНИЯ ПО УМОЛЧАНИЮ
# =====================================================================================
$Defaults = @{
    Bucket      = 'my-bucket'   # имя бакета; можно и ссылку s3://my-bucket/папка
    AwsProfile  = ''            # пусто = профиль AWS CLI по умолчанию
    EndpointUrl = ''            # для S3-совместимых хранилищ
    LogDir      = (Join-Path $PSScriptRoot 'logs')
    LogRetentionDays = 14       # логи этого скрипта старше N дней удаляются при запуске; 0 = не удалять
}
# =====================================================================================

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

$script:AwsProfile = if ($PSBoundParameters.ContainsKey('AwsProfile')) { $AwsProfile } else { $Defaults.AwsProfile }
$script:EndpointUrl = if ($PSBoundParameters.ContainsKey('EndpointUrl')) { $EndpointUrl } else { $Defaults.EndpointUrl }
$script:ExitCode = 0
$script:PrefixBound = $PSBoundParameters.ContainsKey('Prefix')

#region ---------- Вывод и ввод ----------

function Write-Section([string]$Text) { Write-Host ''; Write-Host "=== $Text ===" -ForegroundColor Cyan }
function Write-Info([string]$Text) { Write-Host "  $Text" }
function Write-Ok([string]$Text) { Write-Host "  [OK] $Text" -ForegroundColor Green }
function Write-Warn([string]$Text) { Write-Host "  [!] $Text" -ForegroundColor Yellow }
function Write-Err([string]$Text) { Write-Host "  [X] $Text" -ForegroundColor Red }

# Таблица через Out-String: не зависит от ширины консоли и попадает в лог
function Write-Table { param([Parameter(ValueFromPipeline)]$InputObject) end { $input | Format-Table -AutoSize | Out-String -Width 300 | Write-Host } }

function Format-Size([double]$Bytes) {
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N1} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Format-Date($Value) {
    if ($Value -is [datetime]) { return $Value.ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
    "$Value"
}

# "1,3,5-7" / "all" -> список номеров
function ConvertFrom-SelectionString {
    param([string]$Text, [int]$Count)
    $t = $Text.Trim().ToLower()
    if ($t -in 'all', '*') { return 1..$Count }
    $set = [System.Collections.Generic.SortedSet[int]]::new()
    foreach ($part in ($t -split '[,;\s]+' | Where-Object { $_ })) {
        if ($part -match '^(\d+)-(\d+)$') {
            $a = [int]$Matches[1]; $b = [int]$Matches[2]
            if ($a -gt $b) { $a, $b = $b, $a }
            foreach ($i in $a..$b) { [void]$set.Add($i) }
        }
        elseif ($part -match '^\d+$') { [void]$set.Add([int]$part) }
        else { throw "Cannot parse '$part'." }
    }
    foreach ($i in $set) { if ($i -lt 1 -or $i -gt $Count) { throw "Number $i is out of range 1..$Count." } }
    @($set)
}

#endregion

#region ---------- AWS CLI ----------

function Get-AwsExe {
    $c = Get-Command aws -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $c) { throw 'AWS CLI (aws) not found. Install AWS CLI v2 and run aws configure.' }
    $c.Source
}

# Разбор того, что ввёл пользователь: имя бакета, bucket/папка, s3://, https-ссылка
# (в т.ч. скопированная из консоли AWS) или ARN. Возвращает бакет, папку, подсказку и проблему.
function ConvertFrom-S3Location {
    param([string]$Text)
    $t = "$Text".Trim().Trim('"', "'", '<', '>').Trim()
    $bucket = ''; $prefix = ''; $hint = $null

    if ($t -match '^s3://([^/]*)/?(.*)$') {
        $bucket = $Matches[1]; $prefix = $Matches[2]
    }
    elseif ($t -match '^arn:aws[a-z-]*:s3:::([^/]*)/?(.*)$') {
        $bucket = $Matches[1]; $prefix = $Matches[2]
    }
    elseif ($t -match '^https?://') {
        try { $uri = [uri]$t } catch { return [pscustomobject]@{ Bucket = ''; Prefix = ''; Hint = $null; Problem = "'$t' is not a valid link." } }
        $hostName = $uri.Host.ToLowerInvariant()
        $path = [uri]::UnescapeDataString($uri.AbsolutePath).TrimStart('/')
        $query = @{}
        foreach ($pair in $uri.Query.TrimStart('?') -split '&' | Where-Object { $_ }) {
            $kv = $pair -split '=', 2
            $query[$kv[0]] = if ($kv.Count -gt 1) { [uri]::UnescapeDataString($kv[1].Replace('+', ' ')) } else { '' }
        }
        if ($hostName -like '*console.aws.amazon.com') {
            # https://s3.console.aws.amazon.com/s3/buckets/<bucket>?region=...&prefix=<папка>/
            if ($path -match '^s3/(?:buckets|object)/([^/?]+)') {
                $bucket = $Matches[1]
                if ($query.ContainsKey('prefix')) { $prefix = $query['prefix'] }
            }
        }
        elseif ($hostName -match '^s3([.-][a-z0-9-]+)*\.amazonaws\.com(\.cn)?$') {
            # path-style: https://s3.<region>.amazonaws.com/<bucket>/<папка>
            $parts = $path -split '/', 2
            $bucket = $parts[0]; if ($parts.Count -gt 1) { $prefix = $parts[1] }
        }
        elseif ($hostName -match '^(.+?)\.s3([.-][a-z0-9-]+)*\.amazonaws\.com(\.cn)?$') {
            # virtual-hosted: https://<bucket>.s3.<region>.amazonaws.com/<папка>
            $bucket = $Matches[1]; $prefix = $path
        }
        else {
            # Не AWS: скорее всего S3-совместимое хранилище со ссылкой вида https://<хост>/<bucket>/<папка>
            $parts = $path -split '/', 2
            $bucket = $parts[0]; if ($parts.Count -gt 1) { $prefix = $parts[1] }
            $hint = "'$hostName' is not an AWS address. If this is S3-compatible storage, set EndpointUrl to '$($uri.Scheme)://$($uri.Authority)' (in `$Defaults or with -EndpointUrl)."
        }
    }
    else {
        # «my-bucket» или «my-bucket/папка»
        $parts = $t -split '/', 2
        $bucket = $parts[0]; if ($parts.Count -gt 1) { $prefix = $parts[1] }
    }

    $prefix = "$prefix".Trim('/')
    $problem = $null
    if (-not $bucket) { $problem = "Could not find a bucket name in '$t'." }
    elseif ($bucket -notmatch '^[a-z0-9][a-z0-9.-]{1,61}[a-z0-9]$') {
        # старые бакеты (до 2018 г., us-east-1) могли содержать заглавные буквы и подчёркивания
        if ($bucket -notmatch '^[A-Za-z0-9._-]{3,255}$') {
            $problem = "'$bucket' is not a valid bucket name: only letters, digits, dots and hyphens, 3-63 characters, no spaces."
        }
    }
    [pscustomobject]@{ Bucket = $bucket; Prefix = $prefix; Hint = $hint; Problem = $problem }
}

function Show-S3LocationExamples {
    Write-Info 'Accepted formats:'
    Write-Info '  my-bucket'
    Write-Info '  my-bucket/folder'
    Write-Info '  s3://my-bucket/folder/'
    Write-Info '  https://my-bucket.s3.eu-central-1.amazonaws.com/folder/'
    Write-Info '  https://s3.console.aws.amazon.com/s3/buckets/my-bucket?region=eu-central-1&prefix=folder/'
}

# Человеческое объяснение типичных ошибок aws
function Get-AwsErrorHint([string]$Message) {
    switch -Regex ($Message) {
        'Unable to locate credentials|NoCredentialProviders|could not be found.*profile|config profile' { 'AWS CLI has no credentials for this profile: run "aws configure" (or "aws configure --profile <name>").'; break }
        'InvalidAccessKeyId|SignatureDoesNotMatch' { 'The access key or secret key is wrong: run "aws configure" again.'; break }
        'ExpiredToken|token.*expired|SSO.*expired' { 'The session has expired: log in again (e.g. "aws sso login").'; break }
        '\(404\)|NoSuchBucket|Not Found' { 'No bucket with this name: check the spelling. Enter only the name, s3://name or a link.'; break }
        '\(403\)|Forbidden|AccessDenied' { 'The bucket exists, but this AWS user/profile has no access to it (or it belongs to another account).'; break }
        '\(400\)|Bad Request|\(301\)|PermanentRedirect|AuthorizationHeaderMalformed' { 'The bucket is probably in another region: "aws configure set region <region>" (the region is shown in the AWS console).'; break }
        'Could not connect to the endpoint|Name or service not known|getaddrinfo|timed out|ConnectTimeout|SSL' { 'Network problem: check internet/proxy/VPN, and EndpointUrl if the storage is not AWS.'; break }
        default { $null }
    }
}

# Запуск aws с JSON-выводом; ошибка aws превращается в исключение с её текстом
function Invoke-AwsJson {
    param([string[]]$Arguments)
    $all = @($Arguments) + @('--output', 'json')
    if ($script:AwsProfile) { $all += @('--profile', $script:AwsProfile) }
    if ($script:EndpointUrl) { $all += @('--endpoint-url', $script:EndpointUrl) }
    $out = & (Get-AwsExe) @all 2>&1
    $code = $LASTEXITCODE
    $stderr = ($out | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } | ForEach-Object { "$_" }) -join ' '
    $stdout = ($out | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }) -join "`n"
    if ($code -ne 0) { throw "aws $($Arguments[0..1] -join ' ') failed (exit code $code): $stderr" }
    if ([string]::IsNullOrWhiteSpace($stdout)) { return $null }
    $stdout | ConvertFrom-Json
}

# Папки первого уровня внутри префикса (как строки PRE в aws s3 ls)
function Get-S3Folders {
    param([string]$Bucket, [string]$Prefix)
    $a = @('s3api', 'list-objects-v2', '--bucket', $Bucket, '--delimiter', '/')
    if ($Prefix) { $a += @('--prefix', $Prefix) }
    $r = Invoke-AwsJson $a
    if ($r -and $r.PSObject.Properties['CommonPrefixes'] -and $r.CommonPrefixes) {
        return @($r.CommonPrefixes | ForEach-Object { $_.Prefix } | Sort-Object)
    }
    @()
}

# Все объекты внутри папки, рекурсивно (aws сам докачивает страницы по 1000 объектов)
function Get-S3Objects {
    param([string]$Bucket, [string]$FolderPrefix)
    $r = Invoke-AwsJson @('s3api', 'list-objects-v2', '--bucket', $Bucket, '--prefix', $FolderPrefix)
    if ($r -and $r.PSObject.Properties['Contents'] -and $r.Contents) {
        return @($r.Contents | Sort-Object Key | ForEach-Object {
                [pscustomobject]@{ Key = $_.Key; Size = [int64]$_.Size; LastModified = $_.LastModified }
            })
    }
    @()
}

# Удаление ровно указанных ключей пачками по 1000 (лимит DeleteObjects)
function Remove-S3Keys {
    param([string]$Bucket, [string[]]$Keys, [string]$Activity)
    $deleted = 0
    $errors = [System.Collections.Generic.List[object]]::new()
    $tmp = Join-Path ([IO.Path]::GetTempPath()) "s3-delete-$([guid]::NewGuid()).json"
    try {
        for ($i = 0; $i -lt $Keys.Count; $i += 1000) {
            $batch = @($Keys[$i..([math]::Min($i + 999, $Keys.Count - 1))])
            Write-Progress -Activity $Activity -Status "$i of $($Keys.Count)" -PercentComplete ([int](100 * $i / $Keys.Count))
            $body = @{ Objects = @($batch | ForEach-Object { @{ Key = $_ } }); Quiet = $true } | ConvertTo-Json -Depth 4 -Compress
            [IO.File]::WriteAllText($tmp, $body, [Text.UTF8Encoding]::new($false))
            $r = Invoke-AwsJson @('s3api', 'delete-objects', '--bucket', $Bucket, '--delete', "file://$tmp")
            $batchErrors = @()
            if ($r -and $r.PSObject.Properties['Errors'] -and $r.Errors) { $batchErrors = @($r.Errors) }
            foreach ($e in $batchErrors) { $errors.Add($e) }
            $deleted += $batch.Count - $batchErrors.Count
        }
    }
    finally {
        Write-Progress -Activity $Activity -Completed
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
    [pscustomobject]@{ Deleted = $deleted; Errors = @($errors) }
}

#endregion

#region ---------- Основной сценарий ----------

function Invoke-Main {
    if ($DryRun) { Write-Host 'DRY-RUN MODE: nothing will be deleted.' -ForegroundColor Magenta }

    # ---------- Бакет ----------
    Write-Section 'AWS S3 bucket'
    [void](Get-AwsExe)
    # Бакет можно ввести именем, s3://-ссылкой или ссылкой из консоли AWS
    $raw = $Bucket
    $fromParam = [bool]$Bucket
    if (-not $raw) {
        Write-Info 'Enter the bucket name or a link (s3://bucket/folder, https://..., link from the AWS console).'
        $answer = Read-Host "  Bucket [$($Defaults.Bucket)]"
        $raw = if ([string]::IsNullOrWhiteSpace($answer)) { $Defaults.Bucket } else { $answer }
    }
    while ($true) {
        $loc = ConvertFrom-S3Location $raw
        $problem = $loc.Problem
        if (-not $problem) {
            if ($loc.Hint) { Write-Warn $loc.Hint }
            try { Invoke-AwsJson @('s3api', 'head-bucket', '--bucket', $loc.Bucket) | Out-Null }
            catch {
                $problem = "Bucket '$($loc.Bucket)' is not accessible: $($_.Exception.Message)"
                $h = Get-AwsErrorHint $_.Exception.Message
                if ($h) { $problem += " HINT: $h" }
            }
        }
        if (-not $problem) { break }
        if ($fromParam) { throw $problem }
        Write-Warn $problem
        Show-S3LocationExamples
        $answer = Read-Host '  Bucket (Enter = cancel)'
        if ([string]::IsNullOrWhiteSpace($answer)) { Write-Info 'Cancelled. Nothing was deleted.'; return }
        $raw = $answer
    }
    $bucketName = $loc.Bucket
    Write-Ok "Bucket: $bucketName"

    # Папка из ссылки используется, если -Prefix не передан явно
    $parentRaw = $Prefix
    if ($loc.Prefix -and -not $script:PrefixBound) { $parentRaw = $loc.Prefix }
    $parent = $parentRaw.Trim().Trim('/')
    if ($parent) { $parent += '/' }
    $parentUri = "s3://$bucketName/$parent"

    # ---------- Список папок ----------
    $folders = @(Get-S3Folders -Bucket $bucketName -Prefix $parent)
    if (-not $folders.Count) {
        Write-Info "No folders in $parentUri"
        return
    }
    Write-Info "Folders in ${parentUri}:"
    $i = 0
    $folders | ForEach-Object { $i++; [pscustomobject]@{ '#' = $i; Folder = $_.Substring($parent.Length) } } | Write-Table

    $selected = $null
    while ($null -eq $selected) {
        $answer = Read-Host '  Folders to delete (numbers: 1,3,5-7 or all; Enter = cancel)'
        if ([string]::IsNullOrWhiteSpace($answer)) { Write-Info 'Cancelled. Nothing was deleted.'; return }
        try { $selected = @(ConvertFrom-SelectionString -Text $answer -Count $folders.Count | ForEach-Object { $folders[$_ - 1] }) }
        catch { Write-Warn $_.Exception.Message }
    }

    # ---------- Содержимое выбранных папок ----------
    $plan = @()
    foreach ($folder in $selected) {
        $objects = @(Get-S3Objects -Bucket $bucketName -FolderPrefix $folder)
        $size = ($objects | Measure-Object Size -Sum).Sum
        Write-Section "s3://$bucketName/$folder - $($objects.Count) object(s), $(Format-Size $size)"
        if ($objects.Count) {
            $objects | ForEach-Object {
                $rel = $_.Key.Substring($folder.Length)
                [pscustomobject]@{
                    Object   = if ($rel) { $rel } else { '(folder marker object)' }
                    Size     = Format-Size $_.Size
                    Modified = Format-Date $_.LastModified
                }
            } | Write-Table
        }
        else { Write-Info 'The folder is already empty.' }
        $plan += [pscustomobject]@{ Folder = $folder; Objects = $objects; Size = [int64]$size }
    }

    $totalObjects = ($plan | ForEach-Object { $_.Objects.Count } | Measure-Object -Sum).Sum
    $totalSize = ($plan | Measure-Object Size -Sum).Sum
    if (-not $totalObjects) { Write-Info 'Nothing to delete.'; return }

    # ---------- Итог перед удалением ----------
    Write-Section 'Will be deleted'
    $plan | ForEach-Object { [pscustomobject]@{ Folder = "s3://$bucketName/$($_.Folder)"; Objects = $_.Objects.Count; Size = Format-Size $_.Size } } | Write-Table
    Write-Info "Total: $totalObjects object(s), $(Format-Size $totalSize) in $($plan.Count) folder(s)."

    try {
        $versioning = Invoke-AwsJson @('s3api', 'get-bucket-versioning', '--bucket', $bucketName)
        if ($versioning -and $versioning.PSObject.Properties['Status'] -and $versioning.Status -in 'Enabled', 'Suspended') {
            Write-Warn "Versioning is $($versioning.Status) on this bucket: deleted objects stay as previous versions and still take space (and cost money) until a lifecycle rule removes them."
        }
    }
    catch { Write-Warn "Could not check bucket versioning: $($_.Exception.Message)" }

    if ($DryRun) { Write-Host '  [DRY-RUN] Nothing was deleted.' -ForegroundColor Magenta; return }

    $answer = Read-Host "  Type DELETE to delete these $totalObjects object(s), anything else cancels"
    if ($answer -cne 'DELETE') { Write-Info 'Cancelled. Nothing was deleted.'; return }

    # ---------- Удаление ----------
    $results = @()
    foreach ($item in $plan) {
        if (-not $item.Objects.Count) { continue }
        $uri = "s3://$bucketName/$($item.Folder)"
        Write-Section "Deleting $uri"
        $r = Remove-S3Keys -Bucket $bucketName -Keys @($item.Objects.Key) -Activity "Deleting $uri"
        foreach ($e in $r.Errors) { Write-Err "$($e.Key): $($e.Code) $($e.Message)" }

        # Проверка: что осталось в папке после удаления
        $left = @(Get-S3Objects -Bucket $bucketName -FolderPrefix $item.Folder)
        $planned = [System.Collections.Generic.HashSet[string]]::new([string[]]@($item.Objects.Key))
        $newObjects = @($left | Where-Object { -not $planned.Contains($_.Key) })
        if ($newObjects.Count) {
            Write-Warn "Objects that appeared after the folder was listed were NOT deleted:"
            $newObjects | ForEach-Object { Write-Info "  $($_.Key)" }
        }
        if ($r.Errors.Count -eq 0 -and $left.Count -eq 0) { Write-Ok "Deleted $($r.Deleted) object(s), the folder is gone." }

        $results += [pscustomobject]@{ Folder = $uri; Deleted = $r.Deleted; Errors = $r.Errors.Count; 'Left in folder' = $left.Count }
        if ($r.Errors.Count -or $left.Count) { $script:ExitCode = 1 }
    }

    Write-Section 'Summary'
    $results | Write-Table
}

#endregion

# ---------- Точка входа ----------
$transcript = $false
$oldFileEncoding = $env:AWS_CLI_FILE_ENCODING
try {
    # AWS CLI читает file:// в кодировке системы; ключи с кириллицей требуют UTF-8
    $env:AWS_CLI_FILE_ENCODING = 'UTF-8'
    # Старые логи этого скрипта (только s3-delete-*.log, чужие файлы не трогаем)
    if ($Defaults.LogRetentionDays -gt 0 -and -not $DryRun) {
        $cutoff = (Get-Date).AddDays(-$Defaults.LogRetentionDays)
        Get-ChildItem -LiteralPath $Defaults.LogDir -Filter 's3-delete-*.log' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
    try {
        New-Item -ItemType Directory -Path $Defaults.LogDir -Force | Out-Null
        Start-Transcript -Path (Join-Path $Defaults.LogDir "s3-delete-$(Get-Date -Format 'yyyyMMdd-HHmmss').log") -ErrorAction Stop | Out-Null
        $transcript = $true
    }
    catch { Write-Warn "Logging disabled: $($_.Exception.Message)" }
    Invoke-Main
}
catch {
    Write-Err $_.Exception.Message
    $script:ExitCode = 2
}
finally {
    $env:AWS_CLI_FILE_ENCODING = $oldFileEncoding
    if ($transcript) { try { Stop-Transcript | Out-Null } catch { } }
}
exit $script:ExitCode
