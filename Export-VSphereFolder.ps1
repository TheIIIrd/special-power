#Requires -Version 7.0
<#
.SYNOPSIS
    Exports VMs and templates from a vSphere folder to OVA, AWS S3 and/or VMware Cloud Director.

.DESCRIPTION
    Workflow:
      1. Connect to vCenter and find the folder (if several folders share the name, pick one by full path).
      2. Show the folder contents (templates/VMs, power state, disk size, snapshots, attached ISO images).
      3. Choose the machines and the mode: OVA only / OVA + S3 / OVA + Cloud Director / OVA + S3 + Cloud Director /
         direct vSphere -> Cloud Director transfer without an intermediate OVA (experimental).
      4. Pre-flight checks: S3 and Cloud Director access, name conflicts, free disk space.
      5. One machine at a time: template -> VM, disconnect ISO/floppy, export, back to template, uploads.
      6. At the end ALL powered-off objects in the folder are converted to templates (in a finally block,
         so this also happens on errors and Ctrl+C). A state file is written before any change; if the
         window was closed, run the script with -Restore <state.json>.

    Default values live in $Defaults below. In interactive mode they are offered in the prompt
    (press Enter to accept). Command-line parameters take precedence and are not asked for.

.EXAMPLE
    .\Export-VSphereFolder.ps1
    Fully interactive run.

.EXAMPLE
    .\Export-VSphereFolder.ps1 -DryRun
    Run all checks and show the plan without changing anything.

.EXAMPLE
    .\Export-VSphereFolder.ps1 -NonInteractive -VIServer vc.corp.local -VICredentialFile .\vc.cred.xml `
        -Folder 'DC1/vm/Polygons/P7' -Mode S3AndVcd -S3Prefix 'p7' -VcdCredentialFile .\vcd.cred.xml `
        -OnConflict Suffix -OnPoweredOn Skip -LocalOva Remove

.EXAMPLE
    .\Export-VSphereFolder.ps1 -Mode Vcd -VcdTemplateName @{ 'web01' = 'LAB7 Web Server (2026)'; 'db01' = 'LAB7-DB' }
    Upload to Cloud Director with custom template names for some machines (the rest keep their VM names).

.EXAMPLE
    .\Export-VSphereFolder.ps1 -Mode Vcd -VcdNamePattern 'P7_{name}_{date}'
    Name all templates by a pattern: web01 -> P7_web01_2026-09-17.

.EXAMPLE
    .\Export-VSphereFolder.ps1 -LogRetentionDays 30
    Keep logs for 30 days instead of the default 14 (0 disables cleanup).

.EXAMPLE
    .\Export-VSphereFolder.ps1 -Restore .\logs\state-20260916-101500.json
    Convert the folder objects back to templates after an abnormal termination.

.NOTES
    Credential file for -NonInteractive (DPAPI-encrypted, readable only by the same user on the same machine):
        Get-Credential | Export-Clixml .\vc.cred.xml
#>
[CmdletBinding()]
param(
    [string]$VIServer,
    [pscredential]$VICredential,
    [string]$VICredentialFile,
    [switch]$IgnoreInvalidCertificate,

    [string]$Folder,
    [string[]]$VMName,
    [switch]$Recurse,

    [ValidateSet('ExportOnly', 'S3', 'Vcd', 'S3AndVcd', 'VcdDirect')]
    [string]$Mode,

    [string]$ExportPath,

    [string]$S3Bucket,
    [string]$S3Prefix,
    [string]$AwsProfile,

    [string]$VcdHost,
    [string]$VcdOrg,
    [string]$VcdVdc,
    [string]$VcdCatalog,
    [hashtable]$VcdTemplateName,     # свои имена шаблонов: @{ 'web01' = 'LAB7 Web Server' }
    [string]$VcdNamePattern,         # шаблон имени: '{name}' = имя VM, '{date}' = сегодняшняя дата
    [pscredential]$VcdCredential,
    [string]$VcdCredentialFile,

    [string]$OvfToolPath,

    [ValidateSet('Ask', 'Skip', 'Overwrite', 'Suffix')]
    [string]$OnConflict,

    [ValidateSet('Ask', 'Shutdown', 'Skip', 'Abort')]
    [string]$OnPoweredOn,

    [ValidateSet('Ask', 'Keep', 'Remove')]
    [string]$LocalOva,

    [ValidateRange(0, 3650)]
    [int]$LogRetentionDays,

    [switch]$NonInteractive,
    [switch]$DryRun,
    [string]$Restore
)

# =====================================================================================
#  ЗНАЧЕНИЯ ПО УМОЛЧАНИЮ - поправьте один раз под свою инфраструктуру
# =====================================================================================
$Defaults = [ordered]@{
    VIServer                = 'vcenter.example.local'
    ExportPath              = 'D:\OVA'
    OvfToolPath             = 'C:\Program Files\VMware\VMware OVF Tool\ovftool.exe'

    S3Bucket                = 'my-bucket'     # имя бакета; можно и ссылку s3://my-bucket/папка
    S3Prefix                = ''
    AwsProfile              = ''            # пусто = профиль AWS CLI по умолчанию
    AwsEndpointUrl          = ''            # для S3-совместимых хранилищ, напр. 'https://s3.company.local'
    AwsExtraCopyArgs        = @()           # напр. @('--storage-class', 'STANDARD_IA')

    VcdHost                 = 'vcd.example.com:443'
    VcdOrg                  = 'my-org'
    VcdVdc                  = 'my-vdc'
    VcdCatalog              = 'my-catalog'
    VcdNamePattern          = '{name}'      # имя шаблона в каталоге; напр. 'P7_{name}_{date}'
    VcdUser                 = ''            # подставляется в окно ввода учётки
    VcdApiVersion           = ''            # пусто = определить автоматически
    VcdSkipCertificateCheck = $false

    # Доп. аргументы ovftool, напр. '--noSSLVerify', '--net:VM Network=org-net-01'
    OvfToolExtraArgs        = @('--acceptAllEulas')

    OnConflict              = 'Ask'         # Ask | Skip | Overwrite | Suffix
    OnPoweredOn             = 'Ask'         # Ask | Shutdown | Skip | Abort
    LocalOva                = 'Ask'         # Ask | Keep | Remove
    ShutdownTimeoutSec      = 300
    SpaceReserveFactor      = 1.2           # запас к размеру дисков при проверке места
    LogDir                  = (Join-Path $PSScriptRoot 'logs')
    LogRetentionDays        = 14            # логи старше N дней удаляются при запуске; 0 = не удалять
}
# =====================================================================================

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

$script:BoundParams     = $PSBoundParameters
$script:Stamp           = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:VI              = $null
$script:VICred          = $null
$script:OwnVIConnection = $false
$script:LabFolder       = $null      # объект папки полигона: все поиски VM/темплейтов идут внутри неё
$script:RecurseSearch         = $Recurse.IsPresent
$script:ExitCode        = 0
$script:Cfg             = @{
    LogDir         = $Defaults.LogDir
    AwsProfile     = $Defaults.AwsProfile
    AwsEndpointUrl = $Defaults.AwsEndpointUrl
    VcdApiVersion  = $Defaults.VcdApiVersion
    RemoveOva      = $false
}

#region ---------- Вывод и ввод ----------

function Write-Section([string]$Text) { Write-Host ''; Write-Host "=== $Text ===" -ForegroundColor Cyan }
function Write-Info([string]$Text) { Write-Host "  $Text" }
function Write-Ok([string]$Text) { Write-Host "  [OK] $Text" -ForegroundColor Green }
function Write-Warn([string]$Text) { Write-Host "  [!] $Text" -ForegroundColor Yellow }
function Write-Err([string]$Text) { Write-Host "  [X] $Text" -ForegroundColor Red }

function Read-WithDefault {
    param([string]$Prompt, [string]$Default, [switch]$AllowEmpty)
    while ($true) {
        $hint = ''
        if (-not [string]::IsNullOrEmpty($Default)) { $hint = " [$Default]" }
        if ($AllowEmpty) { $hint += ' (- = empty)' }
        $answer = Read-Host "$Prompt$hint"
        if ($AllowEmpty -and $answer.Trim() -eq '-') { return '' }
        if ([string]::IsNullOrWhiteSpace($answer)) {
            if (-not [string]::IsNullOrEmpty($Default)) { return $Default }
            if ($AllowEmpty) { return '' }
            Write-Warn 'A value is required.'
            continue
        }
        return $answer.Trim()
    }
}

function Read-Choice {
    param([string]$Prompt, [System.Collections.Specialized.OrderedDictionary]$Options, [string]$Default)
    if ($NonInteractive) {
        if ($Default) { return $Default }
        throw "A choice is required ($Prompt), but -NonInteractive is set."
    }
    Write-Host "  $Prompt"
    $keys = @($Options.Keys)
    for ($i = 0; $i -lt $keys.Count; $i++) {
        $mark = if ($keys[$i] -eq $Default) { '  (default)' } else { '' }
        Write-Host ('    {0}) {1}{2}' -f ($i + 1), $Options[$keys[$i]], $mark)
    }
    while ($true) {
        $a = Read-Host '  Choice'
        if ([string]::IsNullOrWhiteSpace($a) -and $Default) { return $Default }
        $n = 0
        if ([int]::TryParse($a, [ref]$n) -and $n -ge 1 -and $n -le $keys.Count) { return $keys[$n - 1] }
        Write-Warn 'Enter a number from the list.'
    }
}

function Confirm-Action {
    param([string]$Prompt, [bool]$Default = $false)
    if ($NonInteractive) { return $Default }
    $hint = if ($Default) { '[Y/n]' } else { '[y/N]' }
    while ($true) {
        $a = "$(Read-Host "  $Prompt $hint")".Trim().ToLower()
        if (-not $a) { return $Default }
        if ($a -in 'y', 'yes', 'д', 'да') { return $true }
        if ($a -in 'n', 'no', 'н', 'нет') { return $false }
    }
}

# Значение настройки: параметр командной строки -> (NonInteractive: $Defaults) -> вопрос с подсказкой из $Defaults.
function Resolve-Setting {
    param([string]$Name, [string]$Prompt, [switch]$AllowEmpty)
    if ($script:BoundParams.ContainsKey($Name)) { return $script:BoundParams[$Name] }
    $default = [string]$Defaults[$Name]
    if ($NonInteractive) {
        if (-not $AllowEmpty -and [string]::IsNullOrWhiteSpace($default)) {
            throw "Parameter -$Name is not set (required with -NonInteractive)."
        }
        return $default
    }
    Read-WithDefault -Prompt "  $Prompt" -Default $default -AllowEmpty:$AllowEmpty
}

# То же, но с проверкой. $Validate возвращает текст ошибки или ничего.
function Resolve-ValidSetting {
    param([string]$Name, [string]$Prompt, [scriptblock]$Validate, [switch]$AllowEmpty)
    $value = Resolve-Setting -Name $Name -Prompt $Prompt -AllowEmpty:$AllowEmpty
    while ($true) {
        $problem = & $Validate $value
        if (-not $problem) { return $value }
        Write-Warn $problem
        if ($NonInteractive) { throw "${Name}: $problem" }
        $value = Read-WithDefault -Prompt "  $Prompt" -Default $value -AllowEmpty:$AllowEmpty
    }
}

function Get-StoredCredential {
    param([pscredential]$Credential, [string]$File, [string]$Message, [string]$UserName)
    if ($Credential) { return $Credential }
    if ($File) {
        $c = Import-Clixml -LiteralPath $File
        if ($c -isnot [pscredential]) { throw "File '$File' does not contain a PSCredential." }
        return $c
    }
    if ($NonInteractive) { throw "No credentials: $Message. Pass a *Credential or *CredentialFile parameter." }
    $p = @{ Message = $Message }
    if ($UserName) { $p.UserName = $UserName }
    $c = Get-Credential @p
    if (-not $c) { throw 'Credential prompt was cancelled.' }
    $c
}

# Разбор "all" / "1,3,5-7" в список индексов.
function ConvertFrom-SelectionString {
    param([string]$Text, [int]$Count)
    $t = "$Text".Trim().ToLower()
    if ($t -in 'all', '*', 'все', 'всё') { return 1..$Count }
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
    if ($set.Count -eq 0) { throw 'Nothing selected.' }
    foreach ($i in $set) { if ($i -lt 1 -or $i -gt $Count) { throw "Number $i is out of range 1..$Count." } }
    return @($set)
}

function Invoke-Step {
    param([Parameter(Mandatory)][string]$Description, [Parameter(Mandatory)][scriptblock]$Action)
    if ($DryRun) { Write-Host "  [DRY-RUN] $Description" -ForegroundColor DarkGray; return }
    Write-Info $Description
    & $Action
}

function Add-Note($Row, [string]$Text) {
    $Row.Note = (@($Row.Note, $Text) | Where-Object { $_ }) -join ' | '
}

#endregion

#region ---------- Утилиты ----------

function Get-SafeFileName([string]$Name) {
    $n = $Name -replace '[<>:"/\\|?*\x00-\x1F]', '_'
    $n.TrimEnd('.', ' ')
}

function Join-S3Key([string]$Prefix, [string]$Name) {
    (@("$Prefix".Trim('/'), $Name) | Where-Object { $_ }) -join '/'
}

# Экранирование аргумента по правилам командной строки Windows.
function Format-NativeArgument([string]$Arg) {
    if ($Arg -eq '') { return '""' }
    if ($Arg -notmatch '[\s"]') { return $Arg }
    $escaped = ($Arg -replace '(\\*)"', '$1$1\"') -replace '(\\+)$', '$1$1'
    '"' + $escaped + '"'
}

# Запуск внешней программы с выводом прямо в консоль (прогресс ovftool/aws виден), возвращает код выхода.
function Start-NativeProcess {
    param([string]$FilePath, [string[]]$ArgumentList)
    $argString = ($ArgumentList | ForEach-Object { Format-NativeArgument $_ }) -join ' '
    $p = Start-Process -FilePath $FilePath -ArgumentList $argString -NoNewWindow -Wait -PassThru
    $p.ExitCode
}

function Protect-Secrets([string]$Text, [string[]]$Secrets) {
    foreach ($s in ($Secrets | Where-Object { $_ } | Sort-Object Length -Descending)) { $Text = $Text.Replace($s, '****') }
    $Text
}

function Get-FreeSpaceGB([string]$Path) {
    try {
        $full = [IO.Path]::GetFullPath($Path)
        $root = [IO.Path]::GetPathRoot($full)
        if (-not $root -or $root.StartsWith('\\')) { return $null }
        [math]::Round([IO.DriveInfo]::new($root).AvailableFreeSpace / 1GB, 1)
    }
    catch { $null }
}

# Удаление старых файлов из папки логов. Трогаем только файлы, которые создаёт сам скрипт
# (на случай, если LogDir указывает на общую папку). Незавершённые state-файлы не удаляем
# независимо от возраста: по ним ещё может понадобиться -Restore.
function Remove-OldLogs {
    param([string]$LogDir, [int]$RetentionDays, [string[]]$Exclude = @())
    if ($RetentionDays -le 0 -or -not (Test-Path -LiteralPath $LogDir -PathType Container)) { return }
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    $patterns = 'run-*.log', 'report-*.csv', 'ovftool-*.log', 'state-*.json'
    $excluded = @($Exclude | Where-Object { $_ } | ForEach-Object { [IO.Path]::GetFullPath($_) })
    $old = @(Get-ChildItem -LiteralPath $LogDir -File |
        Where-Object { $name = $_.Name; @($patterns | Where-Object { $name -like $_ }).Count -gt 0 } |
        Where-Object { $_.LastWriteTime -lt $cutoff -and $_.FullName -notin $excluded })
    if (-not $old.Count) { return }

    $toDelete = @(); $kept = @(); $unreadable = @()
    foreach ($f in $old) {
        if ($f.Name -like 'state-*.json') {
            try { $completed = [bool](Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json).Completed }
            catch { $unreadable += $f; continue }
            if (-not $completed) { $kept += $f; continue }
        }
        $toDelete += $f
    }

    $sizeMB = [math]::Round(($toDelete | Measure-Object Length -Sum).Sum / 1MB, 1)
    if ($toDelete.Count) {
        if ($DryRun) {
            Write-Host "  [DRY-RUN] would delete $($toDelete.Count) log file(s) older than $RetentionDays days ($sizeMB MB)" -ForegroundColor DarkGray
        }
        else {
            $failed = 0
            foreach ($f in $toDelete) {
                try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop } catch { $failed++ }
            }
            Write-Info "Log cleanup: deleted $($toDelete.Count - $failed) file(s) older than $RetentionDays days ($sizeMB MB)."
            if ($failed) { Write-Warn "Could not delete $failed file(s) (in use or no permissions)." }
        }
    }
    foreach ($f in $kept) {
        Write-Warn "Kept old incomplete state file (objects may not be restored yet): $($f.Name). Run -Restore with it; after that it will be cleaned up."
    }
    foreach ($f in $unreadable) {
        Write-Warn "Kept old state file that cannot be read: $($f.Name). Check the folder manually and delete the file."
    }
}

function Save-State([string]$Path, $State) {
    $State | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding utf8
}

#endregion

#region ---------- vSphere ----------

function Import-PowerCLI {
    if (Get-Command Connect-VIServer -ErrorAction SilentlyContinue) { return }
    foreach ($m in 'VCF.PowerCLI', 'VMware.PowerCLI', 'VMware.VimAutomation.Core') {
        if (Get-Module -ListAvailable -Name $m) {
            Import-Module $m -ErrorAction Stop
            if (Get-Command Connect-VIServer -ErrorAction SilentlyContinue) { return }
        }
    }
    throw 'PowerCLI not found. Install it: Install-Module VCF.PowerCLI -Scope CurrentUser'
}

function Connect-VCenter {
    param([string]$Server)
    $existing = @(Get-Variable -Name DefaultVIServers -Scope Global -ValueOnly -ErrorAction SilentlyContinue |
        Where-Object { $_ -and $_.Name -eq $Server -and $_.IsConnected })
    if ($existing) {
        $script:VI = $existing[0]
        Write-Ok "Using existing connection to $Server ($($script:VI.User))"
        return
    }
    if ($IgnoreInvalidCertificate) {
        Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope Session -Confirm:$false | Out-Null
    }
    $credFromParams = $VICredential -or $VICredentialFile
    for ($attempt = 1; ; $attempt++) {
        $script:VICred = Get-StoredCredential -Credential $VICredential -File $VICredentialFile -Message "vCenter $Server credentials"
        try {
            $script:VI = Connect-VIServer -Server $Server -Credential $script:VICred -ErrorAction Stop
            break
        }
        catch {
            if ($NonInteractive -or $credFromParams -or $attempt -ge 3) { throw }
            Write-Warn "Connection failed: $($_.Exception.Message)"
        }
    }
    $script:OwnVIConnection = $true
    Write-Ok "Connected to $Server as $($script:VI.User)"
}

# Путь вида DC1/vm/Полигоны/P7 (корневая папка Datacenters не включается) - тот же формат, что у ovftool vi://.
function Get-VIFolderPath {
    param($Folder)
    $parts = [System.Collections.Generic.List[string]]::new()
    $cur = $Folder
    for ($i = 0; $cur -and $i -lt 64; $i++) {
        $next = $null
        foreach ($p in 'Parent', 'ParentFolder') {
            if ($cur.PSObject.Properties[$p] -and $cur.$p) { $next = $cur.$p; break }
        }
        if (-not $next) {
            foreach ($p in 'ParentId', 'ParentFolderId') {
                if ($cur.PSObject.Properties[$p] -and $cur.$p) {
                    $next = Get-Inventory -Server $script:VI -Id $cur.$p -ErrorAction SilentlyContinue
                    break
                }
            }
        }
        if (-not $next) { break }   # $cur - корневая папка
        $parts.Insert(0, $cur.Name)
        $cur = $next
    }
    $parts -join '/'
}

function Resolve-VMFolder {
    param([string]$Spec)
    $norm = $Spec.Trim().Trim('/')
    $leaf = ($norm -split '/')[-1]
    $found = @(Get-Folder -Server $script:VI -Type VM -Name $leaf -ErrorAction SilentlyContinue |
        ForEach-Object { [pscustomobject]@{ Folder = $_; Path = (Get-VIFolderPath $_) } })
    if ($norm.Contains('/')) {
        $found = @($found | Where-Object {
                $full = $_.Path
                $short = $full -replace '/vm(?=/|$)', ''
                $full -eq $norm -or $short -eq $norm -or $full -like "*/$norm" -or $short -like "*/$norm"
            })
    }
    if ($found.Count -eq 0) { return $null }
    if ($found.Count -eq 1) { return $found[0].Folder }
    if ($NonInteractive) {
        throw "Several folders named '$leaf' found: $($found.Path -join '; '). Pass the full path in -Folder."
    }
    Write-Warn "Several folders named '$leaf' found:"
    $opts = [ordered]@{}
    for ($i = 0; $i -lt $found.Count; $i++) { $opts["$i"] = $found[$i].Path }
    $idx = Read-Choice -Prompt 'Which one?' -Options $opts
    $found[[int]$idx].Folder
}

function New-InventoryItem {
    param($Object, [string]$Kind)
    $ed = $Object.ExtensionData
    $committed = $null
    $media = @()
    $snap = $false
    try { $committed = $ed.Summary.Storage.Committed } catch { }
    try { $snap = [bool]$ed.Snapshot } catch { }
    try {
        $media = @($ed.Config.Hardware.Device |
            Where-Object { $_.Backing -and $_.Backing.GetType().Name -match 'IsoBackingInfo|FloppyImageBackingInfo' } |
            ForEach-Object { $_.Backing.FileName })
    }
    catch { }
    [pscustomobject]@{
        Index           = 0
        Name            = $Object.Name
        Kind            = $Kind
        PowerState      = if ($Kind -eq 'VM') { [string]$Object.PowerState } else { '' }
        SizeGB          = if ($committed) { [math]::Round($committed / 1GB, 1) } else { $null }
        HasSnapshots    = $snap
        Media           = ($media -join '; ')
        FileBase        = ''
        OvaPath         = ''
        LocalAction     = 'None'     # Export | Reuse | Skip | None
        S3Key           = ''
        S3Action        = 'None'     # Upload | Overwrite | Skip | None
        VcdName         = ''
        VcdAction       = 'None'     # Upload | Overwrite | Skip | None
        VcdExistingHref = ''
        PowerAction     = 'None'     # Shutdown | Skip | None
        NeedsExport     = $false
        NeedsDirect     = $false
        NeedsVM         = $false
    }
}

# Поиск объекта в папке полигона по имени — так же, как в ручных командах:
#   Get-Template -Location <папка> -Name <имя> / Get-VM -Location <папка> -Name <имя>
# Id вручную не собираем: формат Id у PowerCLI зависит от типа объекта и версии.
function Find-FolderObject {
    param([string]$Name, [ValidateSet('Template', 'VM', 'Any')][string]$Kind = 'Any')
    $p = @{ Server = $script:VI; Location = $script:LabFolder; ErrorAction = 'SilentlyContinue' }
    if (-not $script:RecurseSearch) { $p.NoRecursion = $true }
    # Имена с [ ] * ? нельзя передать в -Name (это wildcard), тогда берём всё и фильтруем точным сравнением
    if (-not [WildcardPattern]::ContainsWildcardCharacters($Name)) { $p.Name = $Name }
    if ($Kind -in 'Template', 'Any') {
        $t = @(Get-Template @p | Where-Object { $_.Name -eq $Name })
        if ($t.Count) { return [pscustomobject]@{ Kind = 'Template'; Object = $t[0] } }
    }
    if ($Kind -in 'VM', 'Any') {
        $v = @(Get-VM @p | Where-Object { $_.Name -eq $Name })
        if ($v.Count) { return [pscustomobject]@{ Kind = 'VM'; Object = $v[0] } }
    }
    $null
}

function Get-FolderVM([string]$Name) {
    $f = Find-FolderObject -Name $Name -Kind VM
    if (-not $f) { throw "VM '$Name' not found in folder $(Get-VIFolderPath $script:LabFolder)." }
    $f.Object
}

function Get-FolderInventory {
    param($Folder)
    $p = @{ Server = $script:VI; Location = $Folder; ErrorAction = 'Stop' }
    if (-not $script:RecurseSearch) { $p.NoRecursion = $true }
    $items = @()
    $items += @(Get-Template @p | ForEach-Object { New-InventoryItem -Object $_ -Kind 'Template' })
    $items += @(Get-VM @p | ForEach-Object { New-InventoryItem -Object $_ -Kind 'VM' })
    $items = @($items | Sort-Object Name)
    $dup = @($items | Group-Object Name | Where-Object Count -GT 1 | ForEach-Object Name)
    if ($dup.Count) {
        throw "Duplicate names (in different subfolders): $($dup -join ', '). Run without -Recurse or point -Folder at the subfolder."
    }
    for ($i = 0; $i -lt $items.Count; $i++) { $items[$i].Index = $i + 1 }
    $items
}

function Show-Inventory {
    param([object[]]$Items)
    $Items | Format-Table -AutoSize -Wrap -Property `
    @{ n = '#'; e = { $_.Index } },
    @{ n = 'Type'; e = { $_.Kind } },
    @{ n = 'Name'; e = { $_.Name } },
    @{ n = 'Power'; e = { $_.PowerState } },
    @{ n = 'Disks, GB'; e = { $_.SizeGB } },
    @{ n = 'Snapshots'; e = { if ($_.HasSnapshots) { 'yes' } else { '' } } },
    @{ n = 'Attached images'; e = { $_.Media } } | Out-String -Width 300 | Write-Host
    $t = @($Items | Where-Object Kind -EQ 'Template').Count
    Write-Info "Templates: $t, VMs: $($Items.Count - $t)"
    if (@($Items | Where-Object HasSnapshots).Count) {
        Write-Warn 'Some machines have snapshots: snapshots are not included in OVA, only the current disk state is exported.'
    }
}

function Select-InventoryItems {
    param([object[]]$Inventory)
    if ($script:BoundParams.ContainsKey('VMName')) {
        foreach ($pattern in $VMName) {
            if (-not @($Inventory | Where-Object { $_.Name -like $pattern }).Count) { Write-Warn "Pattern '$pattern' matched nothing." }
        }
        $sel = @($Inventory | Where-Object { $n = $_.Name; @($VMName | Where-Object { $n -like $_ }).Count -gt 0 })
        if (-not $sel.Count) { throw '-VMName did not match any machine.' }
        return $sel
    }
    if ($NonInteractive) { return $Inventory }
    while ($true) {
        $answer = Read-WithDefault -Prompt '  Which ones to export? (all or numbers: 1,3,5-7)' -Default 'all'
        try {
            $idx = @(ConvertFrom-SelectionString -Text $answer -Count $Inventory.Count)
            return @($Inventory | Where-Object { $_.Index -in $idx })
        }
        catch { Write-Warn $_.Exception.Message }
    }
}

function Wait-PowerOff([string]$Name, [int]$TimeoutSec) {
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 5
        if ((Get-FolderVM $Name).PowerState -eq 'PoweredOff') { return $true }
    }
    $false
}

function Stop-GuestAndWait {
    param([string]$Name)
    $vm = Get-FolderVM $Name
    if ($vm.PowerState -eq 'PoweredOff') { return $true }
    $timeout = [int]$Defaults.ShutdownTimeoutSec
    $toolsRunning = $vm.PowerState -eq 'PoweredOn' -and $vm.ExtensionData.Guest.ToolsRunningStatus -eq 'guestToolsRunning'
    if ($toolsRunning) {
        Write-Info "Shutting down guest OS of '$Name' (waiting up to $timeout s)..."
        Shutdown-VMGuest -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null
        if (Wait-PowerOff -Name $Name -TimeoutSec $timeout) { Write-Ok 'Powered off.'; return $true }
        Write-Warn "'$Name' did not power off within $timeout s."
    }
    else {
        Write-Warn "'$Name': VMware Tools not running or VM is $($vm.PowerState); graceful shutdown is not possible."
    }
    if (-not (Confirm-Action "Hard power off '$Name' (Stop-VM, like pulling the plug)?" $false)) { return $false }
    Stop-VM -VM (Get-FolderVM $Name) -Confirm:$false -ErrorAction Stop | Out-Null
    Wait-PowerOff -Name $Name -TimeoutSec 120
}

function Clear-VMMedia {
    param($VM)
    $cd = @(Get-CDDrive -VM $VM -ErrorAction Stop | Where-Object { $_.IsoPath -or $_.HostDevice -or $_.RemoteDevice })
    if ($cd.Count) { $cd | Set-CDDrive -NoMedia -Confirm:$false -ErrorAction Stop | Out-Null }
    $fl = @(Get-FloppyDrive -VM $VM -ErrorAction SilentlyContinue | Where-Object { $_.FloppyImagePath -or $_.HostDevice -or $_.RemoteDevice })
    if ($fl.Count) { $fl | Set-FloppyDrive -NoMedia -Confirm:$false -ErrorAction Stop | Out-Null }
}

# Приводит объект к темплейту. Возвращает текстовый статус.
function ConvertTo-TemplateSafe {
    param([string]$Name, [int]$Retries = 3)
    $found = Find-FolderObject -Name $Name
    if (-not $found) { return 'Error: not found in folder' }
    if ($found.Kind -eq 'Template') { return 'Template' }
    $vm = $found.Object
    if ($vm.PowerState -ne 'PoweredOff') { return "Left as VM ($($vm.PowerState))" }
    if ($DryRun) { return 'will become Template (dry-run)' }
    for ($i = 1; $i -le $Retries; $i++) {
        try {
            Write-Info "Converting '$Name' to template"
            Set-VM -VM $vm -ToTemplate -Confirm:$false -ErrorAction Stop | Out-Null
            return 'Template'
        }
        catch {
            if ($i -eq $Retries) { return "Error: $($_.Exception.Message)" }
            Write-Warn "Failed ($($_.Exception.Message)), retrying in 10 s..."
            Start-Sleep -Seconds 10
            $vm = Get-FolderVM $Name
        }
    }
}

function Restore-Templates {
    param([object[]]$Entries)
    foreach ($e in $Entries) {
        [pscustomobject]@{ Name = $e.Name; Result = (ConvertTo-TemplateSafe -Name $e.Name) }
    }
}

function Test-IsTemplateError([string]$Result) { $Result -like 'Error*' }

#endregion

#region ---------- AWS S3 ----------

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

function Get-AwsCommonArgs {
    $a = @()
    if ($script:Cfg.AwsProfile) { $a += @('--profile', $script:Cfg.AwsProfile) }
    if ($script:Cfg.AwsEndpointUrl) { $a += @('--endpoint-url', $script:Cfg.AwsEndpointUrl) }
    $a
}

function Invoke-AwsJson {
    param([string[]]$Arguments)
    $aws = Get-AwsExe
    $all = @($Arguments) + @('--output', 'json') + @(Get-AwsCommonArgs)
    $out = & $aws @all 2>&1
    $code = $LASTEXITCODE
    $stderr = ($out | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } | ForEach-Object { "$_" }) -join ' '
    $stdout = ($out | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }) -join "`n"
    if ($code -ne 0) { throw "aws $($Arguments[0..1] -join ' ') (exit code $code): $stderr" }
    if ([string]::IsNullOrWhiteSpace($stdout)) { return $null }
    $stdout | ConvertFrom-Json
}

function Get-S3Listing {
    param([string]$Bucket, [string]$Prefix)
    $p = "$Prefix".Trim('/')
    $a = @('s3api', 'list-objects-v2', '--bucket', $Bucket, '--delimiter', '/')
    if ($p) { $a += @('--prefix', "$p/") }
    $r = Invoke-AwsJson $a
    $objects = @(); $folders = @()
    if ($r) {
        $skip = if ($p) { $p.Length + 1 } else { 0 }
        if ($r.PSObject.Properties['Contents'] -and $r.Contents) {
            $objects = @($r.Contents | Where-Object { $_.Key.Length -gt $skip } | ForEach-Object {
                    [pscustomobject]@{ Name = $_.Key.Substring($skip); SizeGB = [math]::Round($_.Size / 1GB, 2); LastModified = $_.LastModified }
                })
        }
        if ($r.PSObject.Properties['CommonPrefixes'] -and $r.CommonPrefixes) {
            $folders = @($r.CommonPrefixes | ForEach-Object { $_.Prefix })
        }
    }
    [pscustomobject]@{ Objects = $objects; Folders = $folders }
}

# Вывод содержимого бакета/префикса в виде таблицы — аналог aws s3 ls s3://bucket/prefix/
function Show-S3Listing {
    param([string]$Bucket, [string]$Prefix, $Listing)
    $shown = "s3://$Bucket/$(if ($Prefix) { "$Prefix/" })"
    if (-not $Listing.Objects.Count -and -not $Listing.Folders.Count) {
        if ($Prefix) { Write-Info "$shown is empty or does not exist yet (it will be created on upload)." }
        else { Write-Info "$shown - the bucket is empty." }
        return
    }
    $skip = if ($Prefix) { $Prefix.Length + 1 } else { 0 }
    $rows = @()
    $rows += @($Listing.Folders | Sort-Object | ForEach-Object {
            [pscustomobject]@{ Type = 'DIR'; Name = $_.Substring([math]::Min($skip, $_.Length)); 'Size, GB' = ''; Modified = '' }
        })
    $rows += @($Listing.Objects | Sort-Object Name | ForEach-Object {
            $m = $_.LastModified
            $mod = if ($m -is [datetime]) { $m.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { "$m" }
            [pscustomobject]@{ Type = 'FILE'; Name = $_.Name; 'Size, GB' = $_.SizeGB; Modified = $mod }
        })
    Write-Info "Contents of ${shown}:"
    $rows | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
    Write-Info "Folders: $($Listing.Folders.Count), files: $($Listing.Objects.Count)"
}

function Send-ToS3 {
    param([string]$File, [string]$Bucket, [string]$Key)
    $uri = "s3://$Bucket/$Key"
    if ($DryRun) { Write-Host "  [DRY-RUN] aws s3 cp '$File' $uri" -ForegroundColor DarkGray; return }
    Write-Info "Uploading to $uri"
    $cmdArgs = @('s3', 'cp', $File, $uri) + @($Defaults.AwsExtraCopyArgs) + @(Get-AwsCommonArgs)
    $code = Start-NativeProcess -FilePath (Get-AwsExe) -ArgumentList $cmdArgs
    if ($code -ne 0) { throw "aws s3 cp exited with code $code" }
    $head = Invoke-AwsJson @('s3api', 'head-object', '--bucket', $Bucket, '--key', $Key)
    $local = (Get-Item -LiteralPath $File).Length
    if ([int64]$head.ContentLength -ne $local) {
        throw "S3 object size ($($head.ContentLength)) does not match local file size ($local)"
    }
    Write-Ok "Uploaded and size-verified: $uri"
}

#endregion

#region ---------- Cloud Director ----------

function Get-VcdApiUser {
    $u = $script:Cfg.VcdCredential.UserName
    if ($u -like "*@$($script:Cfg.VcdOrg)") { $u } else { "$u@$($script:Cfg.VcdOrg)" }
}

function Get-VcdOvfUser {
    $u = $script:Cfg.VcdCredential.UserName
    $suffix = "@$($script:Cfg.VcdOrg)"
    if ($u -like "*$suffix") { $u.Substring(0, $u.Length - $suffix.Length) } else { $u }
}

# Каждый раз новая сессия: выгрузка может идти часами, а сессии Cloud Director истекают.
function Connect-VcdApi {
    $c = $script:Cfg
    $base = "https://$($c.VcdHost)"
    $skip = [bool]$Defaults.VcdSkipCertificateCheck
    if (-not $c.VcdApiVersion) {
        $v = Invoke-RestMethod -Uri "$base/api/versions" -SkipCertificateCheck:$skip -ErrorAction Stop
        $list = @($v.SupportedVersions.VersionInfo |
            Where-Object { $_.deprecated -ne 'true' -and "$($_.Version)" -match '^\d+\.\d+$' } |
            ForEach-Object { "$($_.Version)" })
        if (-not $list.Count) { throw 'Could not detect the Cloud Director API version (set VcdApiVersion in $Defaults).' }
        $c.VcdApiVersion = $list | Sort-Object { [version]$_ } | Select-Object -Last 1
    }
    $pair = "$(Get-VcdApiUser):$($c.VcdCredential.GetNetworkCredential().Password)"
    $basic = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
    $resp = Invoke-WebRequest -Uri "$base/cloudapi/1.0.0/sessions" -Method Post -SkipCertificateCheck:$skip -ErrorAction Stop `
        -Headers @{ Authorization = "Basic $basic"; Accept = "application/json;version=$($c.VcdApiVersion)" }
    $token = @($resp.Headers['X-VMWARE-VCLOUD-ACCESS-TOKEN'])[0]
    if (-not $token) { throw 'Cloud Director did not return an access token.' }
    [pscustomobject]@{
        Base     = $base
        SkipCert = $skip
        Headers  = @{ Authorization = "Bearer $token"; Accept = "application/*+json;version=$($c.VcdApiVersion)" }
    }
}

function Get-VcdQuery {
    param($Session, [string]$Type, [string]$Filter)
    $page = 1
    $records = [System.Collections.Generic.List[object]]::new()
    while ($true) {
        $uri = "$($Session.Base)/api/query?type=$Type&format=records&pageSize=100&page=$page&filter=$([uri]::EscapeDataString($Filter))"
        $r = Invoke-RestMethod -Uri $uri -Headers $Session.Headers -SkipCertificateCheck:$Session.SkipCert -ErrorAction Stop
        $batch = @($r.record | Where-Object { $_ })
        foreach ($x in $batch) { $records.Add($x) }
        if ($batch.Count -eq 0 -or $records.Count -ge [int]$r.total) { break }
        $page++
    }
    $records
}

function Wait-VcdTask {
    param($Session, $Task, [int]$TimeoutSec = 1800)
    if (-not $Task -or -not $Task.href) { return }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ($true) {
        $t = Invoke-RestMethod -Uri $Task.href -Headers $Session.Headers -SkipCertificateCheck:$Session.SkipCert -ErrorAction Stop
        if ($t.status -eq 'success') { return }
        if ($t.status -in 'error', 'aborted', 'canceled') { throw "Cloud Director task: $($t.status) $($t.error.message)" }
        if ((Get-Date) -gt $deadline) { throw 'Timed out waiting for the Cloud Director task.' }
        Start-Sleep -Seconds 5
    }
}

function Remove-VcdTemplate {
    param([string]$Href, [string]$Name)
    if ($DryRun) { Write-Host "  [DRY-RUN] delete template '$Name' in Cloud Director" -ForegroundColor DarkGray; return }
    Write-Warn "Deleting existing template '$Name' in Cloud Director..."
    $s = Connect-VcdApi
    $task = Invoke-RestMethod -Method Delete -Uri $Href -Headers $s.Headers -SkipCertificateCheck:$s.SkipCert -ErrorAction Stop
    Wait-VcdTask -Session $s -Task $task
}

function Send-ToVcd {
    param([string]$Source, [string]$TemplateName, [string[]]$ExtraSecrets = @())
    $c = $script:Cfg
    $rawPass = $c.VcdCredential.GetNetworkCredential().Password
    $user = [uri]::EscapeDataString((Get-VcdOvfUser))
    $pass = [uri]::EscapeDataString($rawPass)
    $query = 'org={0}&vdc={1}&catalog={2}&vappTemplate={3}' -f `
    ([uri]::EscapeDataString($c.VcdOrg)), ([uri]::EscapeDataString($c.VcdVdc)),
    ([uri]::EscapeDataString($c.VcdCatalog)), ([uri]::EscapeDataString($TemplateName))
    $target = "vcloud://${user}:${pass}@$($c.VcdHost)/?$query"
    $log = Join-Path $c.LogDir ('ovftool-{0}-{1}.log' -f (Get-SafeFileName $TemplateName), (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $cmdArgs = @("--X:logFile=$log", '--X:logLevel=warning') + @($Defaults.OvfToolExtraArgs) + @($Source, $target)
    $shown = Protect-Secrets -Text ('ovftool ' + (($cmdArgs | ForEach-Object { Format-NativeArgument $_ }) -join ' ')) -Secrets (@($pass, $rawPass) + $ExtraSecrets)
    if ($DryRun) { Write-Host "  [DRY-RUN] $shown" -ForegroundColor DarkGray; return }
    Write-Info $shown
    $code = Start-NativeProcess -FilePath $c.OvfToolPath -ArgumentList $cmdArgs
    if ($code -ne 0) { throw "ovftool exited with code $code (log: $log)" }
    Write-Ok "Template '$TemplateName' uploaded to catalog '$($c.VcdCatalog)'"
}

function Get-ViSourceLocator {
    param($VM)
    $raw = $script:VICred.GetNetworkCredential().Password
    $u = [uri]::EscapeDataString($script:VICred.UserName)
    $p = [uri]::EscapeDataString($raw)
    # Своя папка VM (важно при -Recurse); если свойства нет — папка полигона
    $folderObj = if ($VM.PSObject.Properties['Folder'] -and $VM.Folder) { $VM.Folder } else { $script:LabFolder }
    $path = ((Get-VIFolderPath $folderObj) -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
    [pscustomobject]@{
        Locator = "vi://${u}:${p}@$($script:Cfg.VIServer)/$path/$([uri]::EscapeDataString($VM.Name))"
        Secrets = @($p, $raw)
    }
}

#endregion

#region ---------- План и обработка ----------

# Имя шаблона в каталоге Cloud Director: по шаблону ({name}, {date}), из -VcdTemplateName
# или вручную для каждой машины. Проверяются пустые имена и дубли внутри выборки;
# совпадения с тем, что уже есть в каталоге, проверяются дальше обычным образом.
function Expand-VcdNamePattern([string]$Pattern, [string]$VmName) {
    $Pattern.Replace('{name}', $VmName).Replace('{date}', (Get-Date -Format 'yyyy-MM-dd')).Trim()
}

function Test-VcdTemplateName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'The name is empty.' }
    if ($Name -match '[\x00-\x1F\x7F]') { return 'The name contains control characters.' }
    $null
}

function Resolve-VcdTemplateNames {
    param([object[]]$Items)
    $pattern = if ($script:BoundParams.ContainsKey('VcdNamePattern')) { $VcdNamePattern } else { [string]$Defaults.VcdNamePattern }
    if ([string]::IsNullOrWhiteSpace($pattern)) { $pattern = '{name}' }
    $map = if ($script:BoundParams.ContainsKey('VcdTemplateName') -and $VcdTemplateName) { $VcdTemplateName } else { @{} }
    foreach ($k in @($map.Keys)) {
        if ("$k" -notin $Items.Name) { Write-Warn "-VcdTemplateName: machine '$k' is not in the selection, ignored." }
    }

    Write-Section 'Template names in Cloud Director'
    # Спрашиваем, только если имена не заданы параметрами
    $ask = -not $NonInteractive -and -not $script:BoundParams.ContainsKey('VcdTemplateName') -and -not $script:BoundParams.ContainsKey('VcdNamePattern')
    $choice = 'Same'
    if ($ask) {
        $sameLabel = if ($pattern -eq '{name}') { 'Same as the VM names' } else { "By the pattern from `$Defaults: $pattern" }
        $choice = Read-Choice -Prompt 'How to name the templates?' -Default 'Same' -Options ([ordered]@{
                Same    = $sameLabel
                Pattern = 'By a pattern, e.g. P7_{name}_{date}'
                Each    = 'Enter a name for each machine'
            })
        if ($choice -eq 'Pattern') {
            while ($true) {
                $pattern = Read-WithDefault -Prompt '  Pattern ({name} = VM name, {date} = today)' -Default $pattern
                $problem = Test-VcdTemplateName (Expand-VcdNamePattern $pattern 'x')
                if (-not $problem) { break }
                Write-Warn $problem
            }
        }
    }

    foreach ($it in $Items) {
        $it.VcdName = if ($map.ContainsKey($it.Name)) { "$($map[$it.Name])".Trim() } else { Expand-VcdNamePattern $pattern $it.Name }
        $problem = Test-VcdTemplateName $it.VcdName
        if ($problem) {
            if ($NonInteractive -or -not $ask) { throw "Template name for '$($it.Name)': $problem" }
            $it.VcdName = $it.Name
        }
    }

    if ($choice -eq 'Each') {
        Write-Info 'Enter = keep the suggested name. Spaces and special characters are allowed.'
        foreach ($it in $Items) {
            while ($true) {
                $n = Read-WithDefault -Prompt "  $($it.Name) ->" -Default $it.VcdName
                $problem = Test-VcdTemplateName $n
                if (-not $problem) { $it.VcdName = $n.Trim(); break }
                Write-Warn $problem
            }
        }
    }

    # Две машины не могут загрузиться под одним именем (сравнение без учёта регистра — с запасом)
    while ($true) {
        $dups = @($Items | Group-Object VcdName | Where-Object Count -GT 1)
        if (-not $dups.Count) { break }
        $msg = 'Several machines would get the same template name: ' + (($dups | ForEach-Object { "'$($_.Name)' ($($_.Group.Name -join ', '))" }) -join '; ')
        if ($NonInteractive) { throw $msg }
        Write-Warn $msg
        foreach ($g in $dups) {
            foreach ($it in @($g.Group | Select-Object -Skip 1)) {
                while ($true) {
                    $n = Read-WithDefault -Prompt "  New name for $($it.Name) ->" -Default $it.VcdName
                    $problem = Test-VcdTemplateName $n
                    if (-not $problem) { $it.VcdName = $n.Trim(); break }
                    Write-Warn $problem
                }
            }
        }
    }

    $Items | ForEach-Object { [pscustomobject]@{ VM = $_.Name; 'Template in Cloud Director' = $_.VcdName } } |
        Format-Table -AutoSize | Out-String -Width 300 | Write-Host
}

function Update-ItemPlan {
    param($It, [string]$Mode)
    if ($It.LocalAction -eq 'Skip' -or $It.PowerAction -eq 'Skip') {
        $It.LocalAction = 'Skip'
        if ($It.S3Action -ne 'None') { $It.S3Action = 'Skip' }
        if ($It.VcdAction -ne 'None') { $It.VcdAction = 'Skip' }
    }
    $wantS3 = $It.S3Action -in 'Upload', 'Overwrite'
    $wantVcd = $It.VcdAction -in 'Upload', 'Overwrite'
    $It.NeedsExport = $Mode -ne 'VcdDirect' -and $It.LocalAction -eq 'Export' -and ($Mode -eq 'ExportOnly' -or $wantS3 -or $wantVcd)
    $It.NeedsDirect = $Mode -eq 'VcdDirect' -and $wantVcd
    $It.NeedsVM = $It.NeedsExport -or $It.NeedsDirect
}

function Resolve-ConflictPolicy {
    param([string]$Target, [string]$Suffix)
    $p = if ($script:BoundParams.ContainsKey('OnConflict')) { $OnConflict } else { $Defaults.OnConflict }
    if ($p -ne 'Ask') { return $p }
    if ($NonInteractive) { return 'Skip' }
    Read-Choice -Prompt "How to handle name conflicts in ${Target}?" -Default 'Skip' -Options ([ordered]@{
            Skip      = 'Do not upload these machines there'
            Overwrite = 'Overwrite'
            Suffix    = "Upload with suffix _$Suffix"
        })
}

function Show-Plan {
    param([object[]]$Items)
    $Items | ForEach-Object {
        $it = $_
        [pscustomobject]@{
            'Name'           = $it.Name
            'Type'           = $it.Kind
            'Power'          = if ($it.PowerAction -eq 'Shutdown') { 'SHUT DOWN' } else { $it.PowerState }
            'OVA'            = if ($it.LocalAction -eq 'Skip') { 'skipped' } elseif ($it.NeedsExport) { 'export' } elseif ($it.LocalAction -eq 'Reuse') { 'existing file' } else { '-' }
            'S3'             = switch ($it.S3Action) { 'Upload' { $it.S3Key } 'Overwrite' { "OVERWRITE $($it.S3Key)" } 'Skip' { 'skipped' } default { '-' } }
            'Cloud Director' = switch ($it.VcdAction) {
                'Upload' { if ($it.NeedsDirect) { "direct: $($it.VcdName)" } else { $it.VcdName } }
                'Overwrite' { "OVERWRITE $($it.VcdName)" }
                'Skip' { 'skipped' }
                default { '-' }
            }
        }
    } | Format-Table -AutoSize -Wrap | Out-String -Width 300 | Write-Host
}

function Invoke-ItemPipeline {
    param($Item, [string]$Mode)
    $cfg = $script:Cfg
    $row = [pscustomobject]@{
        Name     = $Item.Name
        Export   = if ($Item.NeedsExport) { 'not done' } elseif ($Item.LocalAction -eq 'Reuse') { 'existing file' } elseif ($Item.LocalAction -eq 'Skip') { 'skipped' } else { '-' }
        S3       = switch ($Item.S3Action) { { $_ -in 'Upload', 'Overwrite' } { 'not done' } 'Skip' { 'skipped' } default { '-' } }
        VCD      = switch ($Item.VcdAction) { { $_ -in 'Upload', 'Overwrite' } { 'not done' } 'Skip' { 'skipped' } default { '-' } }
        Template = ''
        OvaGB    = $null
        Time     = ''
        Note     = ''
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $stage = ''
    $done = if ($DryRun) { 'dry-run' } else { 'OK' }
    try {
        if ($Item.LocalAction -eq 'Skip') { Write-Info 'Skipped.'; return $row }
        $hasUploads = ($Item.S3Action -in 'Upload', 'Overwrite') -or ($Item.VcdAction -in 'Upload', 'Overwrite')
        if (-not $Item.NeedsVM -and -not $hasUploads) { Write-Info 'Nothing to do (everything skipped due to name conflicts).'; return $row }

        if ($Item.NeedsVM) {
            $touched = $false
            try {
                $stage = 'Power'
                if ($Item.Kind -eq 'VM' -and $Item.PowerState -ne 'PoweredOff') {
                    if ($Item.PowerAction -ne 'Shutdown') { throw "VM is $($Item.PowerState)" }
                    if ($DryRun) { Write-Host "  [DRY-RUN] shut down '$($Item.Name)'" -ForegroundColor DarkGray }
                    elseif (-not (Stop-GuestAndWait -Name $Item.Name)) { throw 'VM was not powered off.' }
                }

                $stage = 'Convert to VM'
                $touched = $true
                if ($Item.Kind -eq 'Template') {
                    Invoke-Step 'Template -> VM' {
                        $t = Find-FolderObject -Name $Item.Name -Kind Template
                        if (-not $t) { throw "Template '$($Item.Name)' not found in folder $(Get-VIFolderPath $script:LabFolder)." }
                        Set-Template -Template $t.Object -ToVM -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                }
                $vm = $null
                if (-not $DryRun) { $vm = Get-FolderVM $Item.Name }

                $stage = 'Disconnect media'
                Invoke-Step 'Disconnecting CD/DVD and floppy media' { Clear-VMMedia -VM $vm }

                if ($Item.NeedsExport) {
                    $stage = 'Export'
                    if ($DryRun) {
                        Write-Host "  [DRY-RUN] Export-VApp -> $($Item.OvaPath)" -ForegroundColor DarkGray
                    }
                    else {
                        Write-Info "Export-VApp -> $($Item.OvaPath)"
                        $out = @(Export-VApp -VM $vm -Destination $cfg.ExportPath -Name $Item.FileBase -Format Ova -Force -ErrorAction Stop)
                        $ova = $out | Where-Object { $_ -is [IO.FileInfo] -and $_.Extension -eq '.ova' } | Select-Object -First 1
                        if ($ova) { $Item.OvaPath = $ova.FullName }
                        if (-not (Test-Path -LiteralPath $Item.OvaPath)) { throw "File $($Item.OvaPath) not found after export." }
                    }
                    $row.Export = $done
                }

                if ($Item.NeedsDirect) {
                    $stage = 'Cloud Director'
                    if ($Item.VcdAction -eq 'Overwrite') { Remove-VcdTemplate -Href $Item.VcdExistingHref -Name $Item.VcdName }
                    if ($DryRun) { $src = [pscustomobject]@{ Locator = "vi://<user>@$($cfg.VIServer)/<path>/$($Item.Name)"; Secrets = @() } }
                    else { $src = Get-ViSourceLocator -VM $vm }
                    Send-ToVcd -Source $src.Locator -TemplateName $Item.VcdName -ExtraSecrets $src.Secrets
                    $row.VCD = $done
                }
            }
            finally {
                if ($touched) {
                    $row.Template = ConvertTo-TemplateSafe -Name $Item.Name
                    Write-Info "vSphere: $($row.Template)"
                }
            }
        }

        $needsFile = ($Item.S3Action -in 'Upload', 'Overwrite') -or ($Mode -ne 'VcdDirect' -and $Item.VcdAction -in 'Upload', 'Overwrite')
        if ($needsFile -and -not $DryRun) {
            if (-not (Test-Path -LiteralPath $Item.OvaPath)) { throw "File not found: $($Item.OvaPath)" }
            $row.OvaGB = [math]::Round((Get-Item -LiteralPath $Item.OvaPath).Length / 1GB, 2)
        }

        if ($Item.S3Action -in 'Upload', 'Overwrite') {
            $stage = 'S3'
            try {
                Send-ToS3 -File $Item.OvaPath -Bucket $cfg.S3Bucket -Key $Item.S3Key
                $row.S3 = $done
            }
            catch { $row.S3 = 'Error'; Add-Note $row "S3: $($_.Exception.Message)"; Write-Err "S3: $($_.Exception.Message)" }
        }

        if ($Mode -ne 'VcdDirect' -and $Item.VcdAction -in 'Upload', 'Overwrite') {
            $stage = 'Cloud Director'
            try {
                if ($Item.VcdAction -eq 'Overwrite') { Remove-VcdTemplate -Href $Item.VcdExistingHref -Name $Item.VcdName }
                Send-ToVcd -Source $Item.OvaPath -TemplateName $Item.VcdName
                $row.VCD = $done
            }
            catch { $row.VCD = 'Error'; Add-Note $row "Cloud Director: $($_.Exception.Message)"; Write-Err "Cloud Director: $($_.Exception.Message)" }
        }

        if ($cfg.RemoveOva -and $needsFile) {
            $s3Ok = $Mode -notin 'S3', 'S3AndVcd' -or $row.S3 -eq $done
            $vcdOk = $Mode -notin 'Vcd', 'S3AndVcd' -or $row.VCD -eq $done
            if ($s3Ok -and $vcdOk) {
                Invoke-Step "Deleting local $($Item.OvaPath)" { Remove-Item -LiteralPath $Item.OvaPath -Force }
                if (-not $DryRun) { Add-Note $row 'OVA deleted' }
            }
            else { Add-Note $row 'OVA kept: not all uploads succeeded' }
        }
    }
    catch {
        $msg = $_.Exception.Message
        switch ($stage) {
            'Export' { $row.Export = 'Error' }
            'Cloud Director' { $row.VCD = 'Error' }
            default { $row.Export = 'Error' }
        }
        Add-Note $row "${stage}: $msg"
        Write-Err "${stage}: $msg"
    }
    finally {
        $row.Time = '{0:hh\:mm\:ss}' -f $sw.Elapsed
    }
    $row
}

#endregion

#region ---------- Режим восстановления ----------

function Invoke-RestoreMode {
    $st = Get-Content -LiteralPath $Restore -Raw | ConvertFrom-Json
    $server = if ($script:BoundParams.ContainsKey('VIServer')) { $VIServer } else { $st.VIServer }
    Write-Section "Restoring from $Restore (folder $($st.Folder), created $($st.Created))"
    Connect-VCenter -Server $server
    $script:LabFolder = Resolve-VMFolder -Spec $st.Folder
    if (-not $script:LabFolder) { throw "Folder '$($st.Folder)' from the state file not found." }
    $script:RecurseSearch = [bool]$st.Recurse
    $rows = @(Restore-Templates -Entries @($st.Items))
    $rows | Format-Table Name, Result -AutoSize | Out-String -Width 300 | Write-Host
    if (@($rows | Where-Object { Test-IsTemplateError $_.Result }).Count) {
        $script:ExitCode = 1
    }
    elseif (-not $DryRun) {
        $st.Completed = $true
        Save-State -Path $Restore -State $st
        Write-Ok 'All objects are fine.'
    }
}

#endregion

#region ---------- Основной сценарий ----------

function Invoke-Main {
    $cfg = $script:Cfg
    if ($DryRun) { Write-Host 'DRY-RUN MODE: no changes will be made.' -ForegroundColor Magenta }

    Import-PowerCLI
    if ($Restore) { Invoke-RestoreMode; return }

    # ---------- vCenter ----------
    Write-Section 'Connecting to vSphere'
    $cfg.VIServer = Resolve-Setting -Name VIServer -Prompt 'vCenter'
    Connect-VCenter -Server $cfg.VIServer

    # ---------- Папка ----------
    $folderSpec = if ($script:BoundParams.ContainsKey('Folder')) { $Folder } else { $null }
    while ($true) {
        if (-not $folderSpec) {
            if ($NonInteractive) { throw '-Folder is required with -NonInteractive.' }
            $folderSpec = Read-WithDefault -Prompt '  Lab folder (name or path, e.g. DC1/vm/Polygons/P7)'
        }
        $folderObj = Resolve-VMFolder -Spec $folderSpec
        if (-not $folderObj) {
            if ($NonInteractive) { throw "Folder '$folderSpec' not found." }
            Write-Warn "Folder '$folderSpec' not found."
            $folderSpec = $null; continue
        }
        $script:LabFolder = $folderObj
        $folderPath = Get-VIFolderPath $folderObj
        Write-Section "Contents of $folderPath$(if ($Recurse) { ' (including subfolders)' })"
        $inventory = @(Get-FolderInventory -Folder $folderObj)
        if (-not $inventory.Count) {
            if ($NonInteractive) { throw 'The folder contains no VMs or templates.' }
            Write-Warn 'The folder contains no VMs or templates.'
            $folderSpec = $null; continue
        }
        Show-Inventory $inventory
        if ($NonInteractive -or (Confirm-Action 'Is this the right folder?' $true)) { break }
        $folderSpec = $null
    }

    $selected = @(Select-InventoryItems -Inventory $inventory)
    Write-Info "Selected: $($selected.Count) of $($inventory.Count)"

    # ---------- Режим ----------
    Write-Section 'Mode'
    $mode = if ($script:BoundParams.ContainsKey('Mode')) { $Mode } else {
        Read-Choice -Prompt 'What to do?' -Options ([ordered]@{
                ExportOnly = 'Export to local OVA only'
                S3         = 'OVA + upload to S3'
                Vcd        = 'OVA + upload to Cloud Director'
                S3AndVcd   = 'OVA + S3 + Cloud Director'
                VcdDirect  = 'Direct vSphere -> Cloud Director, no local OVA (experimental)'
            })
    }
    $usesOva = $mode -ne 'VcdDirect'
    $usesS3 = $mode -in 'S3', 'S3AndVcd'
    $usesVcd = $mode -in 'Vcd', 'S3AndVcd', 'VcdDirect'
    Write-Info "Mode: $mode"

    if ($usesOva) {
        $cfg.ExportPath = Resolve-Setting -Name ExportPath -Prompt 'Local folder for OVA files'
        if (-not (Test-Path -LiteralPath $cfg.ExportPath -PathType Container)) {
            if (-not (Confirm-Action "Folder '$($cfg.ExportPath)' does not exist. Create it?" $true)) { throw 'No folder for OVA files.' }
            Invoke-Step "Creating $($cfg.ExportPath)" { New-Item -ItemType Directory -Path $cfg.ExportPath -Force | Out-Null }
        }
    }

    if ($usesVcd) {
        $cfg.OvfToolPath = Resolve-ValidSetting -Name OvfToolPath -Prompt 'Path to ovftool' -Validate {
            param($v) if (-not (Test-Path -LiteralPath $v -PathType Leaf)) { "ovftool not found: $v" }
        }
    }
    if ($mode -eq 'VcdDirect' -and -not $script:VICred) {
        # ovftool нужен логин/пароль vCenter, а подключение PowerCLI было взято уже готовым
        $script:VICred = Get-StoredCredential -Credential $VICredential -File $VICredentialFile -Message "vCenter $($cfg.VIServer) credentials for ovftool"
    }

    # ---------- S3 ----------
    $s3Existing = @()
    if ($usesS3) {
        Write-Section 'AWS S3'
        [void](Get-AwsExe)
        $cfg.AwsProfile = Resolve-Setting -Name AwsProfile -Prompt 'AWS CLI profile' -AllowEmpty
        # Бакет можно ввести именем, s3://-ссылкой или ссылкой из консоли AWS
        $raw = Resolve-Setting -Name S3Bucket -Prompt 'Bucket (name, s3://bucket/folder or a link from the AWS console)'
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
            Write-Warn $problem
            if ($NonInteractive) { throw "S3Bucket: $problem" }
            Show-S3LocationExamples
            $raw = Read-WithDefault -Prompt '  Bucket' -Default $raw
        }
        $cfg.S3Bucket = $loc.Bucket
        if ($loc.Bucket -ne "$raw".Trim()) { Write-Ok "Bucket: $($loc.Bucket)" }
        # Сначала показываем, что уже лежит в бакете, и только потом спрашиваем папку
        $root = Get-S3Listing -Bucket $cfg.S3Bucket -Prefix ''
        Show-S3Listing -Bucket $cfg.S3Bucket -Prefix '' -Listing $root

        if ($loc.Prefix -and -not $script:BoundParams.ContainsKey('S3Prefix')) {
            # Папка пришла из ссылки — предлагаем её по умолчанию
            Write-Info "Folder from the link: $($loc.Prefix)/"
            $prefix = if ($NonInteractive) { $loc.Prefix } else { Read-WithDefault -Prompt '  Folder (prefix) in the bucket for this lab' -Default $loc.Prefix -AllowEmpty }
        }
        else {
            $prefix = Resolve-Setting -Name S3Prefix -Prompt 'Folder (prefix) in the bucket for this lab' -AllowEmpty
        }
        while ($true) {
            $prefix = "$prefix".Trim().Trim('/')
            $listing = if ($prefix) { Get-S3Listing -Bucket $cfg.S3Bucket -Prefix $prefix } else { $root }
            if ($prefix) { Show-S3Listing -Bucket $cfg.S3Bucket -Prefix $prefix -Listing $listing }
            if ($NonInteractive -or $script:BoundParams.ContainsKey('S3Prefix') -or (Confirm-Action 'Upload here?' $true)) { break }
            $prefix = Read-WithDefault -Prompt '  Prefix' -Default $prefix -AllowEmpty
        }
        $cfg.S3Prefix = $prefix
        $s3Existing = @($listing.Objects | ForEach-Object { $_.Name })
    }

    # ---------- Cloud Director ----------
    $vcdExisting = $null
    if ($usesVcd) {
        Write-Section 'VMware Cloud Director'
        $cfg.VcdHost = Resolve-Setting -Name VcdHost -Prompt 'Address (host:port)'
        $cfg.VcdOrg = Resolve-Setting -Name VcdOrg -Prompt 'Organization'
        $cfg.VcdVdc = Resolve-Setting -Name VcdVdc -Prompt 'VDC'
        $cfg.VcdCatalog = Resolve-Setting -Name VcdCatalog -Prompt 'Catalog'
        $cfg.VcdCredential = Get-StoredCredential -Credential $VcdCredential -File $VcdCredentialFile `
            -Message "Cloud Director credentials (organization $($cfg.VcdOrg))" -UserName $Defaults.VcdUser
        try {
            $session = Connect-VcdApi
            Write-Ok "Logged in (API $($cfg.VcdApiVersion))"
            if (-not @(Get-VcdQuery -Session $session -Type 'catalog' -Filter "name==$($cfg.VcdCatalog)").Count) {
                Write-Warn "Catalog '$($cfg.VcdCatalog)' not found or no permissions on it."
                if (-not (Confirm-Action 'Continue anyway?' $false)) { throw 'Catalog not found.' }
            }
            $records = @(Get-VcdQuery -Session $session -Type 'vAppTemplate' -Filter "catalogName==$($cfg.VcdCatalog)")
            $vcdExisting = @{}
            foreach ($r in $records) { $vcdExisting[[string]$r.name] = $r }
            Write-Info "Templates in catalog '$($cfg.VcdCatalog)': $($records.Count)"
            if ($records.Count) { $records | Sort-Object name | Format-Table name, status, creationDate -AutoSize | Out-String -Width 300 | Write-Host }
        }
        catch {
            if ($_.Exception.Message -eq 'Catalog not found.') { throw }
            Write-Warn "Could not check Cloud Director via API: $($_.Exception.Message)"
            if (-not (Confirm-Action 'Continue without name conflict checks (ovftool will fail on duplicates)?' $false)) {
                throw 'Cloud Director is not accessible.'
            }
        }
    }

    # ---------- Начальный план ----------
    $suffix = Get-Date -Format 'yyyyMMdd-HHmm'
    foreach ($it in $selected) {
        $it.FileBase = Get-SafeFileName $it.Name
        if ($usesOva) { $it.OvaPath = Join-Path $cfg.ExportPath "$($it.FileBase).ova"; $it.LocalAction = 'Export' }
        if ($usesS3) { $it.S3Key = Join-S3Key $cfg.S3Prefix "$($it.FileBase).ova"; $it.S3Action = 'Upload' }
        if ($usesVcd) { $it.VcdName = $it.Name; $it.VcdAction = 'Upload' }
    }
    # Свои имена шаблонов — до проверки совпадений с каталогом, чтобы она шла уже по новым именам
    if ($usesVcd) { Resolve-VcdTemplateNames -Items $selected }

    # ---------- Совпадения ----------
    if ($usesOva) {
        $localConf = @($selected | Where-Object { Test-Path -LiteralPath $_.OvaPath })
        if ($localConf.Count) {
            Write-Section 'Local OVA files already exist'
            foreach ($it in $localConf) {
                $f = Get-Item -LiteralPath $it.OvaPath
                Write-Info ('{0}  ({1} GB, {2:yyyy-MM-dd HH:mm})' -f $f.FullName, [math]::Round($f.Length / 1GB, 2), $f.LastWriteTime)
            }
            $choice = Read-Choice -Prompt 'What to do with them?' -Default 'Export' -Options ([ordered]@{
                    Export = 'Export again (overwrite)'
                    Reuse  = 'Use the existing files, do not export from vSphere'
                    Skip   = 'Skip these machines'
                })
            foreach ($it in $localConf) { $it.LocalAction = $choice }
        }
    }

    if ($usesS3) {
        $conf = @($selected | Where-Object { $_.LocalAction -ne 'Skip' -and $s3Existing -ccontains "$($_.FileBase).ova" })
        if ($conf.Count) {
            Write-Section 'Name conflicts in S3'
            $conf | ForEach-Object { Write-Info $_.S3Key }
            $policy = Resolve-ConflictPolicy -Target 'S3' -Suffix $suffix
            foreach ($it in $conf) {
                switch ($policy) {
                    'Skip' { $it.S3Action = 'Skip' }
                    'Overwrite' { $it.S3Action = 'Overwrite' }
                    'Suffix' { $it.S3Key = Join-S3Key $cfg.S3Prefix "$($it.FileBase)_$suffix.ova" }
                }
            }
        }
    }

    if ($usesVcd -and $vcdExisting) {
        $conf = @($selected | Where-Object { $_.LocalAction -ne 'Skip' -and $vcdExisting.ContainsKey($_.VcdName) })
        if ($conf.Count) {
            Write-Section 'Name conflicts in Cloud Director'
            $conf | ForEach-Object { Write-Info $_.VcdName }
            $policy = Resolve-ConflictPolicy -Target 'Cloud Director' -Suffix $suffix
            if ($policy -eq 'Overwrite') { Write-Warn 'The old template is deleted BEFORE the new upload: if the upload fails, it will be gone.' }
            foreach ($it in $conf) {
                switch ($policy) {
                    'Skip' { $it.VcdAction = 'Skip' }
                    'Overwrite' { $it.VcdAction = 'Overwrite'; $it.VcdExistingHref = $vcdExisting[$it.VcdName].href }
                    'Suffix' { $it.VcdName = "$($it.VcdName)_$suffix" }
                }
            }
        }
    }

    foreach ($it in $selected) { Update-ItemPlan -It $it -Mode $mode }

    # ---------- Включённые машины ----------
    $poweredOn = @($selected | Where-Object { $_.NeedsVM -and $_.Kind -eq 'VM' -and $_.PowerState -ne 'PoweredOff' })
    if ($poweredOn.Count) {
        Write-Section 'Powered-on machines'
        $poweredOn | ForEach-Object { Write-Info "$($_.Name) ($($_.PowerState))" }
        $policy = if ($script:BoundParams.ContainsKey('OnPoweredOn')) { $OnPoweredOn } else { $Defaults.OnPoweredOn }
        if ($policy -eq 'Ask') {
            $policy = if ($NonInteractive) { 'Skip' } else {
                Read-Choice -Prompt 'They cannot be exported while running. After export they become templates and will not be powered on again.' -Default 'Shutdown' -Options ([ordered]@{
                        Shutdown = 'Shut down via guest OS (Shutdown-VMGuest)'
                        Skip     = 'Skip them'
                        Abort    = 'Abort the script'
                    })
            }
        }
        if ($policy -eq 'Abort') { throw 'Aborted: the selection contains powered-on VMs.' }
        foreach ($it in $poweredOn) { $it.PowerAction = $policy; Update-ItemPlan -It $it -Mode $mode }
    }

    # ---------- Удаление OVA после загрузки ----------
    if ($usesOva -and ($usesS3 -or $usesVcd)) {
        $p = if ($script:BoundParams.ContainsKey('LocalOva')) { $LocalOva } else { $Defaults.LocalOva }
        $cfg.RemoveOva = switch ($p) {
            'Remove' { $true }
            'Keep' { $false }
            default { Confirm-Action 'Delete the local OVA after a successful upload to all destinations?' $false }
        }
    }

    # ---------- Место на диске ----------
    $toExport = @($selected | Where-Object NeedsExport)
    if ($toExport.Count) {
        $sizes = @($toExport | Where-Object { $null -ne $_.SizeGB } | ForEach-Object { $_.SizeGB })
        $baseGB = if (-not $sizes.Count) { 0 } elseif ($cfg.RemoveOva) { ($sizes | Measure-Object -Maximum).Maximum } else { ($sizes | Measure-Object -Sum).Sum }
        $needGB = [math]::Round($baseGB * $Defaults.SpaceReserveFactor, 1)
        $free = Get-FreeSpaceGB $cfg.ExportPath
        if ($null -eq $free) { Write-Warn 'Could not determine free space (network path?): check skipped.' }
        elseif ($free -lt $needGB) {
            Write-Warn "Free: $free GB, estimated need: ~$needGB GB (OVA is usually smaller than the disks, but not always)."
            if (-not (Confirm-Action 'Continue?' $false)) { throw 'Not enough disk space for OVA files.' }
        }
        else { Write-Ok "Disk space: $free GB free, estimated need ~$needGB GB." }
    }

    # ---------- Итоговый план ----------
    Write-Section 'Plan'
    Show-Plan -Items $selected
    $becomeTemplates = @($inventory | Where-Object { $_.Kind -eq 'VM' -and ($_.PowerState -eq 'PoweredOff' -or $_.PowerAction -eq 'Shutdown') })
    $stayVM = @($inventory | Where-Object { $_.Kind -eq 'VM' -and $_.PowerState -ne 'PoweredOff' -and $_.PowerAction -ne 'Shutdown' })
    if ($becomeTemplates.Count) { Write-Info "These original VMs will also become templates at the end: $($becomeTemplates.Name -join ', ')" }
    if ($stayVM.Count) { Write-Warn "Will stay VMs (powered on, left untouched): $($stayVM.Name -join ', ')" }

    if (-not (Confirm-Action 'Start?' $true)) { Write-Info 'Cancelled.'; return }

    # ---------- Выполнение ----------
    $statePath = Join-Path $cfg.LogDir "state-$($script:Stamp).json"
    $state = [ordered]@{
        Created   = (Get-Date).ToString('s')
        VIServer  = $cfg.VIServer
        Folder    = $folderPath
        Recurse   = $script:RecurseSearch
        Completed = $false
        Items     = @($inventory | ForEach-Object { [ordered]@{ Name = $_.Name; OriginalKind = $_.Kind } })
    }
    if (-not $DryRun) {
        Save-State -Path $statePath -State $state
        Write-Info "State saved: $statePath"
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $final = @()
    try {
        $n = 0
        foreach ($it in $selected) {
            $n++
            Write-Section "[$n/$($selected.Count)] $($it.Name)"
            $results.Add((Invoke-ItemPipeline -Item $it -Mode $mode))
        }
    }
    finally {
        Write-Section 'Converting folder objects to templates'
        $final = @(Restore-Templates -Entries $state.Items)
        $byName = @{}
        foreach ($f in $final) { $byName[$f.Name] = $f.Result }
        foreach ($r in $results) { if ($byName.ContainsKey($r.Name)) { $r.Template = $byName[$r.Name] } }
        $bad = @($final | Where-Object { Test-IsTemplateError $_.Result })
        if (-not $DryRun) {
            $state.Completed = ($bad.Count -eq 0)
            Save-State -Path $statePath -State $state
        }
        if ($bad.Count) {
            Write-Err "Failed to convert back to templates: $($bad.Name -join ', ')"
            Write-Err "Retry: .\$(Split-Path -Leaf $PSCommandPath) -Restore '$statePath'"
        }
        else { Write-Ok 'All powered-off objects in the folder are templates.' }
    }

    # ---------- Отчёт ----------
    Write-Section 'Summary'
    $results | Format-Table -AutoSize -Wrap -Property `
    @{ n = 'Name'; e = { $_.Name } },
    @{ n = 'OVA'; e = { $_.Export } },
    @{ n = 'S3'; e = { $_.S3 } },
    @{ n = 'Cloud Director'; e = { $_.VCD } },
    @{ n = 'vSphere'; e = { $_.Template } },
    @{ n = 'OVA, GB'; e = { $_.OvaGB } },
    @{ n = 'Time'; e = { $_.Time } },
    @{ n = 'Note'; e = { $_.Note } } | Out-String -Width 300 | Write-Host

    $others = @($final | Where-Object { $_.Name -notin $selected.Name })
    if ($others.Count) {
        Write-Info 'Other objects in the folder:'
        $others | Format-Table Name, Result -AutoSize | Out-String -Width 300 | Write-Host
    }

    if (-not $DryRun) {
        $csv = Join-Path $cfg.LogDir "report-$($script:Stamp).csv"
        $results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8BOM
        Write-Info "Report: $csv"
    }

    $failed = @($results | Where-Object { $_.Export -eq 'Error' -or $_.S3 -eq 'Error' -or $_.VCD -eq 'Error' -or (Test-IsTemplateError $_.Template) })
    if ($failed.Count -or @($final | Where-Object { Test-IsTemplateError $_.Result }).Count) { $script:ExitCode = 1 }
}

#endregion

# ---------- Точка входа ----------
$transcript = $false
try {
    New-Item -ItemType Directory -Path $script:Cfg.LogDir -Force | Out-Null
    try {
        Start-Transcript -Path (Join-Path $script:Cfg.LogDir "run-$($script:Stamp).log") -ErrorAction Stop | Out-Null
        $transcript = $true
    }
    catch { Write-Warn "Logging disabled: $($_.Exception.Message)" }

    $retention = if ($script:BoundParams.ContainsKey('LogRetentionDays')) { $LogRetentionDays } else { [int]$Defaults.LogRetentionDays }
    try { Remove-OldLogs -LogDir $script:Cfg.LogDir -RetentionDays $retention -Exclude @($Restore) }
    catch { Write-Warn "Log cleanup failed: $($_.Exception.Message)" }

    Invoke-Main
}
catch {
    Write-Err $_.Exception.Message
    if ($_.InvocationInfo) { Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor DarkGray }
    $script:ExitCode = 2
}
finally {
    if ($script:OwnVIConnection -and $script:VI) {
        try { Disconnect-VIServer -Server $script:VI -Confirm:$false -ErrorAction Stop | Out-Null } catch { }
    }
    if ($transcript) { try { Stop-Transcript | Out-Null } catch { } }
}
exit $script:ExitCode
