<#
.SYNOPSIS
    Creates (and optionally replicates) FlashBlade file system snapshots over SSH,
    then destroys and eradicates policy-less snapshots that have reached their retention.
 
.DESCRIPTION
    Tested against the Purity//FB 4.5.x CLI syntax:
        purefs snapshot send   [--suffix SUFFIX] [--target TARGET] FILE-SYSTEM
        purefs snapshot create [--suffix SUFFIX] FILE-SYSTEM
        purefs snapshot list   --csv
        purefs snapshot destroy   FILE-SYSTEM-SNAPSHOT
        purefs snapshot eradicate FILE-SYSTEM-SNAPSHOT
 
    Phase 1 - Snapshot / replicate
        For every file system in the JSON file:
          * targetFilesystem present -> purefs snapshot send --suffix <SUFFIX> --target <targetArray> <fs>
          * targetFilesystem missing -> purefs snapshot create --suffix <SUFFIX> <fs>   (local only, not replicated)
 
        SUFFIX = <RETENTION>-<timestamp>, e.g. 14D-20261012-153000
        The timestamp is required because a suffix must be unique per file system
        (a second run with only "14D" would fail with "already exists").
 
    Phase 2 - Retention cleanup
        On every source array (and optionally every target array) the script runs
        "purefs snapshot list --csv". Any snapshot that
          * has NO policy, and
          * has a suffix that starts with a retention token (30MIN, 12H, 14D, 5W, 13M), and
          * is as old or older than that retention (based on the "Created" column)
        is destroyed and then eradicated.
 
    Retention units (case-insensitive):
        MIN = minutes, H = hours, D = days, W = weeks, M = months
        Allowed range: 1 minute .. 2 years.
 
    Authentication (per array, resolved before any work starts):
        1. SSH key stored locally (sshKeyPath, or ~/.ssh/id_ed25519 | id_ecdsa | id_rsa)
           -> native OpenSSH client (ssh.exe)
        2. No local key, but a credentialFile saved via Export-Clixml
           -> Posh-SSH with that credential
              Get-Credential pureuser | Export-Clixml C:\Scripts\FBSnapshots\flashblade3.cred.xml
        3. Neither -> the script prompts for the pureuser password (3 tries per array)
           -> Posh-SSH (Install-Module Posh-SSH -Scope CurrentUser)
        authMethod "password" in the JSON skips step 1.
        Use -NonInteractive for scheduled runs: the script then fails instead of waiting at a prompt.
 
.PARAMETER ConfigPath
    Path to the JSON configuration file.
 
.PARAMETER SkipSnapshots
    Only run the retention cleanup phase.
 
.PARAMETER SkipCleanup
    Only run the snapshot / replicate phase.
 
.PARAMETER NonInteractive
    Never prompt for a password. An array with no local SSH key and no saved credential
    causes the run to stop with an error (use this for Task Scheduler).
 
.EXAMPLE
    .\Invoke-FBSnapshotReplication.ps1 -ConfigPath .\fb-snapshot-config.json -WhatIf
    Dry run: shows every command that would be sent, changes nothing.
 
.EXAMPLE
    .\Invoke-FBSnapshotReplication.ps1 -ConfigPath .\fb-snapshot-config.json -Verbose
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'fb-snapshot-config.json'),
 
    [switch]$SkipSnapshots,
    [switch]$SkipCleanup,
    [switch]$NonInteractive
)
 
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
 
#region ---------- Globals ----------
$script:LogFile      = $null
$script:ErrorCount   = 0
$script:SshSessions  = @{}          # Posh-SSH sessions keyed by array name
$script:AuthInfo     = @{}          # resolved login method keyed by array name
$script:Settings     = $null
$script:ArrayTz      = [TimeZoneInfo]::Local
 
$NamePattern         = '^[A-Za-z0-9][A-Za-z0-9_-]{0,62}$'      # file system names
$ArrayNamePattern    = '^[A-Za-z0-9][A-Za-z0-9._-]{0,62}$'     # array / connection names
$SnapNamePattern     = '^[A-Za-z0-9][A-Za-z0-9._-]*$'          # snapshot names returned by the array
$RetentionPattern    = '^(?<num>\d{1,6})(?<unit>MIN|H|D|W|M)$'   # full retention string in JSON
$SuffixRetentionRx   = '^(?<num>\d{1,6})(?<unit>MIN|H|D|W|M)(?:-|$)' # retention token at start of a suffix
#endregion
 
#region ---------- Logging ----------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR' { $script:ErrorCount++; Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'DEBUG' { Write-Verbose $line }
        default { Write-Host $line }
    }
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line -WhatIf:$false }
}
#endregion
 
#region ---------- Retention helpers ----------
function ConvertTo-Retention {
    <# Parses "14D", "5W", "13M", "12H", "30MIN" into an object. Returns $null if not parseable. #>
    param([Parameter(Mandatory)][string]$Text, [switch]$FromSuffix)
 
    $rx = if ($FromSuffix) { $SuffixRetentionRx } else { $RetentionPattern }
    $m  = [regex]::Match($Text.Trim(), $rx, 'IgnoreCase')
    if (-not $m.Success) { return $null }
 
    [pscustomobject]@{
        Value = [int]$m.Groups['num'].Value
        Unit  = $m.Groups['unit'].Value.ToUpperInvariant()
        Text  = ('{0}{1}' -f [int]$m.Groups['num'].Value, $m.Groups['unit'].Value.ToUpperInvariant())
    }
}
 
function Add-Retention {
    <# Returns $Date + retention, using calendar math for months. #>
    param([Parameter(Mandatory)][datetime]$Date, [Parameter(Mandatory)]$Retention)
    switch ($Retention.Unit) {
        'MIN' { return $Date.AddMinutes($Retention.Value) }
        'H'   { return $Date.AddHours($Retention.Value) }
        'D'   { return $Date.AddDays($Retention.Value) }
        'W'   { return $Date.AddDays(7 * $Retention.Value) }
        'M'   { return $Date.AddMonths($Retention.Value) }
    }
    throw "Unknown retention unit '$($Retention.Unit)'"
}
 
function Test-RetentionRange {
    <# 1 minute <= retention <= 2 years #>
    param([Parameter(Mandatory)]$Retention)
    if ($Retention.Value -lt 1) { return $false }
    $ref = [datetime]'2000-01-01'
    return ((Add-Retention -Date $ref -Retention $Retention) -le $ref.AddYears(2))
}
 
function ConvertFrom-ArrayTime {
    <#
        Converts the "Created" column ("2026-10-12 15:30:00 EDT" or "2026-10-12 15:30:00")
        to a DateTime in the array's local time zone (Kind = Unspecified).
    #>
    param([Parameter(Mandatory)][string]$Text)
    $m = [regex]::Match($Text, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}')
    if (-not $m.Success) { return $null }
    $dt = [datetime]::ParseExact($m.Value, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    return [datetime]::SpecifyKind($dt, [DateTimeKind]::Unspecified)
}
 
function ConvertTo-UtcFromArrayTime {
    param([Parameter(Mandatory)][datetime]$ArrayLocal)
    $unspec = [datetime]::SpecifyKind($ArrayLocal, [DateTimeKind]::Unspecified)
    if ($script:ArrayTz.IsInvalidTime($unspec)) { $unspec = $unspec.AddHours(1) }   # DST spring-forward gap
    return [TimeZoneInfo]::ConvertTimeToUtc($unspec, $script:ArrayTz)
}
#endregion
 
#region ---------- SSH ----------
function Get-ConfigValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $Name -and
        $null -ne $Object.$Name -and "$($Object.$Name)".Trim() -ne '') {
        return $Object.$Name
    }
    return $Default
}
 
function Test-CanPrompt {
    <# True when a person can answer a Read-Host prompt in this session. #>
    if ($NonInteractive) { return $false }
    if (-not [Environment]::UserInteractive) { return $false }
    if ([Environment]::GetCommandLineArgs() | Where-Object { $_ -match '^-NonI' }) { return $false }
    return $true
}
 
function Find-LocalSshKey {
    <#
        Returns the path of the SSH private key to use for an array, or $null when no key is stored locally.
        Order: sshKeyPath from JSON, then the default OpenSSH key files in the user's .ssh folder.
    #>
    param([Parameter(Mandatory)]$Array)
 
    $key = Get-ConfigValue $Array 'sshKeyPath'
    if ($key) {
        $key = [Environment]::ExpandEnvironmentVariables($key)
        if (Test-Path -LiteralPath $key -PathType Leaf) { return $key }
        Write-Log -Level WARN -Message "[$($Array.name)] sshKeyPath '$key' was not found on this host."
        return $null
    }
 
    $sshDir = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.ssh'
    foreach ($f in 'id_ed25519', 'id_ecdsa', 'id_rsa') {
        $p = Join-Path $sshDir $f
        if (Test-Path -LiteralPath $p -PathType Leaf) { return $p }
    }
    return $null
}
 
function New-FBPasswordSession {
    <# Opens a Posh-SSH session with a password credential. Returns the session or throws. #>
    param([Parameter(Mandatory)]$Array, [Parameter(Mandatory)][pscredential]$Credential)
 
    if (-not (Get-Module -Name Posh-SSH)) {
        if (-not (Get-Module -ListAvailable -Name Posh-SSH)) {
            throw "Password authentication requires the Posh-SSH module. Install it with: Install-Module Posh-SSH -Scope CurrentUser"
        }
        Import-Module Posh-SSH -ErrorAction Stop
    }
    $connectTimeout = [int](Get-ConfigValue $script:Settings 'sshConnectTimeoutSeconds' 15)
    New-SSHSession -ComputerName $Array.managementAddress -Credential $Credential `
                   -AcceptKey -ConnectionTimeout $connectTimeout -ErrorAction Stop
}
 
function Resolve-FBAuth {
    <#
        Decides how to log in to an array before any work starts:
          * an SSH key stored locally   -> ssh.exe with that key
          * a saved credentialFile      -> Posh-SSH with that credential
          * neither                     -> prompt for the pureuser password (up to 3 tries) -> Posh-SSH
        Password logins are tested immediately so a wrong password is caught up front.
    #>
    param([Parameter(Mandatory)]$Array)
 
    $user       = Get-ConfigValue $script:Settings 'sshUser' 'pureuser'
    $authMethod = (Get-ConfigValue $Array 'authMethod' 'key').ToLowerInvariant()
 
    # 1. SSH key
    if ($authMethod -eq 'key') {
        $key = Find-LocalSshKey -Array $Array
        if ($key) {
            Write-Log "[$($Array.name)] Using SSH key $key"
            return [pscustomobject]@{ Mode = 'key'; KeyPath = $key }
        }
        Write-Log -Level WARN -Message "[$($Array.name)] No SSH key stored locally; password required for $user@$($Array.managementAddress)."
    }
 
    # 2. Saved credential file
    $credFile = Get-ConfigValue $Array 'credentialFile'
    if ($credFile -and (Test-Path -LiteralPath $credFile -PathType Leaf)) {
        $cred = Import-Clixml -Path $credFile
        $cred = New-Object System.Management.Automation.PSCredential($user, $cred.Password)   # always log in as sshUser
        $script:SshSessions[$Array.name] = New-FBPasswordSession -Array $Array -Credential $cred
        Write-Log "[$($Array.name)] Logged in with saved credential $credFile"
        return [pscustomobject]@{ Mode = 'password' }
    }
    elseif ($credFile) {
        Write-Log -Level WARN -Message "[$($Array.name)] credentialFile '$credFile' was not found."
    }
 
    # 3. Ask the user
    if (-not (Test-CanPrompt)) {
        throw "[$($Array.name)] No SSH key or saved credential, and this session cannot prompt for a password (non-interactive)."
    }
 
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $secure = Read-Host -AsSecureString -Prompt "Enter password for $user@$($Array.managementAddress) ($($Array.name))"
        if ($null -eq $secure -or $secure.Length -eq 0) {
            Write-Log -Level WARN -Message "[$($Array.name)] Empty password entered."
            continue
        }
        $cred = New-Object System.Management.Automation.PSCredential($user, $secure)
        try {
            $script:SshSessions[$Array.name] = New-FBPasswordSession -Array $Array -Credential $cred
            Write-Log "[$($Array.name)] Logged in as $user with the entered password."
            return [pscustomobject]@{ Mode = 'password' }
        }
        catch {
            if ($_.Exception.Message -match 'Posh-SSH') { throw }
            Write-Log -Level WARN -Message "[$($Array.name)] Login failed (attempt $attempt of 3): $($_.Exception.Message)"
        }
    }
    throw "[$($Array.name)] Could not log in as $user after 3 attempts."
}
 
function Invoke-FBCommand {
    <#
        Runs a single Purity//FB CLI command on an array.
        Returns [pscustomobject]@{ ExitCode; Output (string[]) }
    #>
    param(
        [Parameter(Mandatory)]$Array,
        [Parameter(Mandatory)][string]$Command
    )
 
    $user           = Get-ConfigValue $script:Settings 'sshUser' 'pureuser'
    $connectTimeout = [int](Get-ConfigValue $script:Settings 'sshConnectTimeoutSeconds' 15)
    $cmdTimeout     = [int](Get-ConfigValue $script:Settings 'sshCommandTimeoutSeconds' 300)
    $auth           = $script:AuthInfo[$Array.name]
    if (-not $auth) { throw "[$($Array.name)] Authentication was not resolved for this array." }
 
    Write-Log -Level DEBUG -Message "[$($Array.name)] ssh $user@$($Array.managementAddress) '$Command'"
 
    if ($auth.Mode -eq 'password') {
        # ---------- Posh-SSH (saved or prompted password) ----------
        $sess = $script:SshSessions[$Array.name]
        $r    = Invoke-SSHCommand -SessionId $sess.SessionId -Command $Command -TimeOut $cmdTimeout -ErrorAction Stop
        $out  = @($r.Output)
        if ($r.Error) { $out += @("$($r.Error)" -split "`r?`n") }
        return [pscustomobject]@{ ExitCode = [int]$r.ExitStatus; Output = $out }
    }
    else {
        # ---------- Native OpenSSH (key based) ----------
        $sshArgs = @(
            '-o', 'BatchMode=yes',
            '-o', 'StrictHostKeyChecking=accept-new',
            '-o', "ConnectTimeout=$connectTimeout",
            '-o', 'ServerAliveInterval=30',
            '-i', $auth.KeyPath,
            "$user@$($Array.managementAddress)",
            $Command
        )
 
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'      # native stderr must not throw in Windows PowerShell 5.1
        try {
            $raw  = & ssh.exe @sshArgs 2>&1
            $code = $LASTEXITCODE
        }
        finally { $ErrorActionPreference = $prevEap }
 
        $out = @($raw | ForEach-Object { "$_" })
        return [pscustomobject]@{ ExitCode = [int]$code; Output = $out }
    }
}
 
function Close-FBSessions {
    foreach ($s in $script:SshSessions.Values) {
        try { Remove-SSHSession -SessionId $s.SessionId | Out-Null } catch { }
    }
    $script:SshSessions.Clear()
}
#endregion
 
#region ---------- Snapshot listing ----------
function Get-FBSnapshot {
    <# Runs "purefs snapshot list --csv" and returns objects: Name, Source, Created, Policies #>
    param([Parameter(Mandatory)]$Array)
 
    $r = Invoke-FBCommand -Array $Array -Command 'purefs snapshot list --csv'
    if ($r.ExitCode -ne 0) {
        throw "[$($Array.name)] 'purefs snapshot list --csv' failed (exit $($r.ExitCode)): $($r.Output -join ' | ')"
    }
 
    $lines = @($r.Output | Where-Object { $_ -and $_.Trim() -ne '' })
    if ($lines.Count -eq 0) { return @() }
 
    if ($lines[0] -match '^\s*"?Name"?\s*,') {
        $rows = $lines | ConvertFrom-Csv
    }
    else {
        $rows = $lines | ConvertFrom-Csv -Header 'Name', 'Source', 'Created', 'Policies'
    }
 
    foreach ($row in $rows) {
        $policy = $null
        if ($row.PSObject.Properties.Name -contains 'Policies') { $policy = $row.Policies }
        elseif ($row.PSObject.Properties.Name -contains 'Policy') { $policy = $row.Policy }
 
        [pscustomobject]@{
            Name     = "$($row.Name)".Trim()
            Source   = "$($row.Source)".Trim()
            Created  = "$($row.Created)".Trim()
            Policies = "$policy".Trim()
        }
    }
}
 
function Test-SourceMatch {
    <# Source is "fs" on the source array and "remote-array:fs" on a replication target #>
    param([string]$Source, [string[]]$FileSystems)
    $plain = ($Source -split ':')[-1]
    return ($FileSystems -contains $plain)
}
#endregion
 
#region ---------- Load & validate config ----------
if (-not (Test-Path $ConfigPath)) { throw "Config file not found: $ConfigPath" }
$config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
$script:Settings = $config.settings
 
$logDir = Get-ConfigValue $script:Settings 'logDirectory' (Join-Path $PSScriptRoot 'Logs')
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null }
$script:LogFile = Join-Path $logDir ('FBSnapshot_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
 
Write-Log "===== FlashBlade snapshot/replication run started (config: $ConfigPath) ====="
if ($WhatIfPreference) { Write-Log -Level WARN -Message 'WhatIf mode: no changes will be made on any array.' }
 
$tzId = Get-ConfigValue $script:Settings 'arrayTimeZone'
if ($tzId) {
    try { $script:ArrayTz = [TimeZoneInfo]::FindSystemTimeZoneById($tzId) }
    catch { throw "settings.arrayTimeZone '$tzId' is not a valid time zone ID on this host." }
}
Write-Log "Array time zone for 'Created' timestamps: $($script:ArrayTz.Id)"
 
$tsFormat         = Get-ConfigValue $script:Settings 'suffixTimestampFormat' 'yyyyMMdd-HHmmss'
$cleanupScope     = Get-ConfigValue $script:Settings 'cleanupScope' 'Configured'
$cleanupTargets   = [bool](Get-ConfigValue $script:Settings 'cleanupTargetArrays' $false)
$keepLatest       = [bool](Get-ConfigValue $script:Settings 'keepLatestSnapshotPerFilesystem' $true)
 
# Arrays lookup
$arrays = @{}
foreach ($a in @($config.arrays)) {
    if ($a.name -notmatch $ArrayNamePattern) { throw "Invalid array name '$($a.name)'." }
    if (-not (Get-ConfigValue $a 'managementAddress')) { throw "Array '$($a.name)' has no managementAddress." }
    $arrays[$a.name] = $a
}
 
# File systems
$fsList = @()
foreach ($fs in @($config.filesystems)) {
    if ($fs.name -notmatch $NamePattern) { throw "Invalid file system name '$($fs.name)'." }
 
    $src = Get-ConfigValue $fs 'sourceArray'
    if (-not $src -or -not $arrays.ContainsKey($src)) {
        throw "File system '$($fs.name)': sourceArray '$src' is not defined in the arrays section."
    }
 
    $ret = ConvertTo-Retention -Text "$(Get-ConfigValue $fs 'retention' '')"
    if (-not $ret)                         { throw "File system '$($fs.name)': invalid retention '$($fs.retention)' (use e.g. 30MIN, 12H, 14D, 5W, 13M)." }
    if (-not (Test-RetentionRange $ret))   { throw "File system '$($fs.name)': retention '$($fs.retention)' must be between 1 minute and 2 years." }
 
    $tgtFs    = Get-ConfigValue $fs 'targetFilesystem'
    $tgtArray = Get-ConfigValue $fs 'targetArray'
    $replicate = [bool]$tgtFs
    if ($replicate) {
        if (-not $tgtArray)                         { throw "File system '$($fs.name)': targetFilesystem is set but targetArray is missing." }
        if ($tgtArray -notmatch $ArrayNamePattern)  { throw "File system '$($fs.name)': invalid targetArray '$tgtArray'." }
        if ($tgtFs -notmatch $NamePattern)          { throw "File system '$($fs.name)': invalid targetFilesystem '$tgtFs'." }
    }
 
    $fsList += [pscustomobject]@{
        Name             = $fs.name
        SourceArray      = $src
        Retention        = $ret
        Replicate        = $replicate
        TargetArray      = $tgtArray
        TargetFilesystem = $tgtFs
    }
}
Write-Log ("Loaded {0} array(s) and {1} file system(s)." -f $arrays.Count, $fsList.Count)
#endregion
 
try {
    #region ---------- Resolve logins (prompts happen here, before any change is made) ----------
    $usedArrays = New-Object System.Collections.Generic.List[string]
    foreach ($fs in $fsList) {
        if (-not $usedArrays.Contains($fs.SourceArray)) { $usedArrays.Add($fs.SourceArray) }
        if ($cleanupTargets -and -not $SkipCleanup -and $fs.Replicate -and $arrays.ContainsKey($fs.TargetArray) -and
            -not $usedArrays.Contains($fs.TargetArray)) { $usedArrays.Add($fs.TargetArray) }
    }
    foreach ($name in $usedArrays) {
        try { $script:AuthInfo[$name] = Resolve-FBAuth -Array $arrays[$name] }
        catch { Write-Log -Level ERROR -Message $_.Exception.Message }
    }
    $unresolved = @($usedArrays | Where-Object { -not $script:AuthInfo.ContainsKey($_) })
    if ($unresolved.Count -gt 0) {
        throw "Cannot log in to: $($unresolved -join ', '). No changes were made."
    }
    #endregion
 
    #region ---------- Phase 1: snapshot + replicate ----------
    if (-not $SkipSnapshots) {
        Write-Log '----- Phase 1: create / send snapshots -----'
        foreach ($fs in $fsList) {
            $array  = $arrays[$fs.SourceArray]
            $suffix = '{0}-{1}' -f $fs.Retention.Text, (Get-Date -Format $tsFormat)
 
            if ($suffix.Length -gt 63 -or $suffix -notmatch '^[A-Za-z0-9][A-Za-z0-9-]*[A-Za-z0-9]$') {
                Write-Log -Level ERROR -Message "[$($array.name)] Suffix '$suffix' is not valid for Purity//FB (alphanumeric and '-', max 63). Check suffixTimestampFormat."
                continue
            }
 
            if ($fs.Replicate) {
                $cmd  = "purefs snapshot send --suffix $suffix --target $($fs.TargetArray) $($fs.Name)"
                $what = "snapshot + replicate $($fs.Name) -> $($fs.TargetArray):$($fs.TargetFilesystem) (suffix $suffix)"
            }
            else {
                $cmd  = "purefs snapshot create --suffix $suffix $($fs.Name)"
                $what = "local snapshot of $($fs.Name) (suffix $suffix, not replicated)"
            }
 
            if ($PSCmdlet.ShouldProcess("$($array.name)", $cmd)) {
                try {
                    $r = Invoke-FBCommand -Array $array -Command $cmd
                    if ($r.ExitCode -eq 0) {
                        Write-Log "[$($array.name)] OK: $what"
                        $r.Output | ForEach-Object { Write-Log -Level DEBUG -Message "    $_" }
                    }
                    else {
                        Write-Log -Level ERROR -Message "[$($array.name)] FAILED (exit $($r.ExitCode)): $cmd :: $($r.Output -join ' | ')"
                    }
                }
                catch {
                    Write-Log -Level ERROR -Message "[$($array.name)] FAILED: $cmd :: $($_.Exception.Message)"
                }
            }
        }
    }
    #endregion
 
    #region ---------- Phase 2: retention cleanup ----------
    if (-not $SkipCleanup) {
        Write-Log '----- Phase 2: retention cleanup -----'
 
        # Which arrays to clean, and which file systems are "managed" on each
        $cleanupMap = @{}   # arrayName -> list of file system names
        foreach ($fs in $fsList) {
            if (-not $cleanupMap.ContainsKey($fs.SourceArray)) { $cleanupMap[$fs.SourceArray] = New-Object System.Collections.Generic.List[string] }
            $cleanupMap[$fs.SourceArray].Add($fs.Name)
 
            if ($cleanupTargets -and $fs.Replicate) {
                if ($arrays.ContainsKey($fs.TargetArray)) {
                    if (-not $cleanupMap.ContainsKey($fs.TargetArray)) { $cleanupMap[$fs.TargetArray] = New-Object System.Collections.Generic.List[string] }
                    $cleanupMap[$fs.TargetArray].Add($fs.TargetFilesystem)
                }
                else {
                    Write-Log -Level WARN -Message "cleanupTargetArrays is on but target '$($fs.TargetArray)' has no entry in 'arrays'; skipping target cleanup for $($fs.TargetFilesystem)."
                }
            }
        }
 
        $nowUtc = [datetime]::UtcNow
 
        foreach ($arrayName in $cleanupMap.Keys) {
            $array   = $arrays[$arrayName]
            $managed = @($cleanupMap[$arrayName] | Select-Object -Unique)
 
            try { $snaps = @(Get-FBSnapshot -Array $array) }
            catch { Write-Log -Level ERROR -Message $_.Exception.Message; continue }
 
            Write-Log "[$arrayName] $($snaps.Count) snapshot(s) returned by 'purefs snapshot list --csv'."
 
            # Newest snapshot per source file system (protected when keepLatestSnapshotPerFilesystem = true)
            $latestBySource = @{}
            foreach ($s in $snaps) {
                $c = ConvertFrom-ArrayTime $s.Created
                if (-not $c) { continue }
                if (-not $latestBySource.ContainsKey($s.Source) -or $c -gt $latestBySource[$s.Source].Created) {
                    $latestBySource[$s.Source] = [pscustomobject]@{ Name = $s.Name; Created = $c }
                }
            }
 
            $expired = @()
            foreach ($s in $snaps) {
                if ($cleanupScope -ne 'All' -and -not (Test-SourceMatch -Source $s.Source -FileSystems $managed)) { continue }
 
                # Policy-managed snapshots are never touched
                if ($s.Policies -ne '' -and $s.Policies -ne '-') { continue }
 
                if ($s.Name -notmatch $SnapNamePattern) {
                    Write-Log -Level WARN -Message "[$arrayName] Skipping snapshot with unexpected name '$($s.Name)'."
                    continue
                }
 
                # Suffix = text after the last '.', e.g. fs1.14D-20261012-153000 or fs1.replica.14D-20261012-153000
                $suffix = ($s.Name -split '\.')[-1]
                $ret    = ConvertTo-Retention -Text $suffix -FromSuffix
                if (-not $ret) {
                    Write-Log -Level DEBUG -Message "[$arrayName] $($s.Name): no retention token in suffix, skipped."
                    continue
                }
                if (-not (Test-RetentionRange $ret)) {
                    Write-Log -Level WARN -Message "[$arrayName] $($s.Name): retention '$($ret.Text)' outside 1 minute - 2 years, skipped."
                    continue
                }
 
                $created = ConvertFrom-ArrayTime $s.Created
                if (-not $created) {
                    Write-Log -Level WARN -Message "[$arrayName] $($s.Name): could not parse Created '$($s.Created)', skipped."
                    continue
                }
 
                $expiresUtc = ConvertTo-UtcFromArrayTime (Add-Retention -Date $created -Retention $ret)
                if ($nowUtc -lt $expiresUtc) {
                    Write-Log -Level DEBUG -Message "[$arrayName] $($s.Name): keeps until $($expiresUtc.ToLocalTime()) (retention $($ret.Text))."
                    continue
                }
 
                if ($keepLatest -and $latestBySource.ContainsKey($s.Source) -and $latestBySource[$s.Source].Name -eq $s.Name) {
                    Write-Log -Level WARN -Message "[$arrayName] $($s.Name): expired but it is the newest snapshot of '$($s.Source)'; kept (keepLatestSnapshotPerFilesystem)."
                    continue
                }
 
                $expired += [pscustomobject]@{ Name = $s.Name; Source = $s.Source; Created = $s.Created; Retention = $ret.Text }
            }
 
            if ($expired.Count -eq 0) {
                Write-Log "[$arrayName] No expired snapshots."
                continue
            }
 
            foreach ($e in $expired) {
                Write-Log "[$arrayName] Expired: $($e.Name) (source $($e.Source), created $($e.Created), retention $($e.Retention))"
 
                $destroyCmd   = "purefs snapshot destroy $($e.Name)"
                $eradicateCmd = "purefs snapshot eradicate $($e.Name)"
 
                if ($PSCmdlet.ShouldProcess("$arrayName : $($e.Name)", 'Destroy and eradicate snapshot')) {
                    try {
                        $r = Invoke-FBCommand -Array $array -Command $destroyCmd
                        if ($r.ExitCode -ne 0) {
                            Write-Log -Level ERROR -Message "[$arrayName] destroy FAILED (exit $($r.ExitCode)): $($e.Name) :: $($r.Output -join ' | ')"
                            continue
                        }
                        Write-Log "[$arrayName] Destroyed $($e.Name)"
 
                        $r = Invoke-FBCommand -Array $array -Command $eradicateCmd
                        if ($r.ExitCode -ne 0) {
                            Write-Log -Level ERROR -Message "[$arrayName] eradicate FAILED (exit $($r.ExitCode)): $($e.Name) :: $($r.Output -join ' | ')"
                            continue
                        }
                        Write-Log "[$arrayName] Eradicated $($e.Name)"
                    }
                    catch {
                        Write-Log -Level ERROR -Message "[$arrayName] cleanup FAILED for $($e.Name) :: $($_.Exception.Message)"
                    }
                }
            }
        }
    }
    #endregion
}
catch {
    Write-Log -Level ERROR -Message $_.Exception.Message
}
finally {
    Close-FBSessions
    Write-Log ("===== Run finished with {0} error(s). Log: {1} =====" -f $script:ErrorCount, $script:LogFile)
}
 
if ($script:ErrorCount -gt 0) { exit 1 } else { exit 0 }
 