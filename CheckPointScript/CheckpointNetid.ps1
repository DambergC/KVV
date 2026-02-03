#TestArash
#Requires -Version 5.1
[CmdletBinding()]
param()

#region ==================== KONFIGURATION ====================

$Config = @{
    LogPath                 = "C:\FileStore\Logs"
    LogPrefix               = "VPNConnect"
    ClearLogEachRun         = $true

    # Check Point
    CheckPointPath          = "C:\Program Files (x86)\CheckPoint\Endpoint Connect"
    TracExe                 = "trac.exe"
    VPNSite                 = "vpn-ao.kriminalvarden.se"

    # Inside-detection
    InternalTestHost        = "kvv.se"
    InternalTestPort        = 636

    # Timing
    CertWaitMaxSeconds      = 180
    CertPollMs              = 300
    VPNConnectTimeoutSec    = 60
    PostConnectWaitSeconds  = 60   # post-check efter connect

    # Cert-urval
    FriendlyIssuerTag       = "ISSUCARSAK0"
    AdminExcludeRegex       = '(?i)(\s-\s(ladm|adm)\b|domadm|utv|test)'

    # UI suppression (best effort)
    UIPatternsToStop        = @("trgui", "cptray", "endpointconnect")
    UIStopMinIntervalSec    = 10

    # Anti-double-run
    LockFile                = "C:\FileStore\VPNConnect.lock"
    LockMaxAgeMinutes       = 3

    # Services (best effort) – för att trac list ska fungera stabilt
    CheckpointServicesStartOrder = @("MADService","TracSrvWrapper","EPWD")
    CheckpointServiceWaitSeconds = 10

    # Net iD
    NetIdPath               = "C:\Program Files\Pointsharp\Net iD\Client\netid.exe"
}

#endregion

#region ==================== LOGGNING ====================

function Initialize-Logging {
    if (-not (Test-Path -Path $Config.LogPath)) {
        $null = New-Item -Path $Config.LogPath -ItemType Directory -Force
    }

    $logFileName = "{0}_{1}.log" -f $Config.LogPrefix, $env:USERNAME
    $script:LogFile = Join-Path -Path $Config.LogPath -ChildPath $logFileName

    if ($Config.ClearLogEachRun) {
        try { Set-Content -Path $script:LogFile -Value "" -Encoding UTF8 } catch {}
    } else {
        if (-not (Test-Path $script:LogFile)) {
            try { Set-Content -Path $script:LogFile -Value "" -Encoding UTF8 } catch {}
        }
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARNING","ERROR","SUCCESS")][string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $entry = "[{0}] [{1}] {2}" -f $ts, $Level, $Message
    try { Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8 } catch {}
}

function Invoke-Step {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][scriptblock]$ScriptBlock
    )
    $t0 = Get-Date
    Write-Log -Message ("---> START: {0}" -f $Name) -Level "INFO"
    try {
        $result = & $ScriptBlock
        $ms = [int]((New-TimeSpan -Start $t0 -End (Get-Date)).TotalMilliseconds)
        Write-Log -Message ("<--- OK: {0} ({1} ms)" -f $Name, $ms) -Level "SUCCESS"
        return $result
    } catch {
        $ms = [int]((New-TimeSpan -Start $t0 -End (Get-Date)).TotalMilliseconds)
        Write-Log -Message ("<--- FAIL: {0} ({1} ms) | {2}" -f $Name, $ms, $_.Exception.Message) -Level "ERROR"
        throw
    }
}

#endregion

#region ==================== LÅS ====================

function Acquire-Lock {
    try {
        if (Test-Path $Config.LockFile) {
            $age  = (Get-Item $Config.LockFile).LastWriteTime
            $mins = (New-TimeSpan -Start $age -End (Get-Date)).TotalMinutes
            if ($mins -lt $Config.LockMaxAgeMinutes) {
                Write-Log -Message ("Lock finns och är {0:N1} min gammal (< {1} min). Avbryter." -f $mins, $Config.LockMaxAgeMinutes) -Level "WARNING"
                return $false
            }
            Write-Log -Message ("Stale lock ({0:N1} min). Tar bort och fortsätter." -f $mins) -Level "WARNING"
            Remove-Item $Config.LockFile -Force -ErrorAction SilentlyContinue
        }

        $null = New-Item -Path $Config.LockFile -ItemType File -ErrorAction Stop
        Set-Content -Path $Config.LockFile -Value ("PID={0}`nUser={1}`nTime={2}" -f $PID, (whoami), (Get-Date)) -Encoding UTF8
        Write-Log -Message ("Lock skapat: {0}" -f $Config.LockFile) -Level "INFO"
        return $true
    } catch {
        Write-Log -Message ("Kunde inte skapa lock: {0}" -f $_.Exception.Message) -Level "ERROR"
        return $false
    }
}

function Release-Lock {
    try {
        if (Test-Path $Config.LockFile) {
            Remove-Item $Config.LockFile -Force -ErrorAction SilentlyContinue
            Write-Log -Message "Lock borttaget." -Level "INFO"
        }
    } catch {}
}

#endregion

#region ==================== NETID ====================

function Start-NetId {
    $netIdPath = $Config.NetIdPath
    if (-not (Test-Path -Path $netIdPath -PathType Leaf)) {
        Write-Log -Message ("netid.exe saknas: {0}" -f $netIdPath) -Level "ERROR"
        throw "netid.exe saknas"
    }

    Write-Log -Message ("Startar Net iD: {0}" -f $netIdPath) -Level "INFO"
    try {
        Start-Process -FilePath $netIdPath -ErrorAction Stop | Out-Null
        Write-Log -Message "Net iD startad." -Level "SUCCESS"
    } catch {
        Write-Log -Message ("Kunde inte starta Net iD: {0}" -f $_.Exception.Message) -Level "ERROR"
        throw
    }
}

#endregion

#region ==================== NÄTVERK ====================

function Test-InsideNetwork {
    Write-Log -Message ("Inside-test: {0}:{1}" -f $Config.InternalTestHost, $Config.InternalTestPort) -Level "INFO"
    try {
        $r = Test-NetConnection -ComputerName $Config.InternalTestHost -Port $Config.InternalTestPort -WarningAction SilentlyContinue
        if ($r.RemoteAddress) { Write-Log -Message ("RemoteAddress={0}" -f $r.RemoteAddress) -Level "INFO" }
        Write-Log -Message ("TcpTestSucceeded={0}" -f $r.TcpTestSucceeded) -Level "INFO"
        return [bool]$r.TcpTestSucceeded
    } catch {
        Write-Log -Message ("Inside-test exception: {0}" -f $_.Exception.Message) -Level "WARNING"
        return $false
    }
}

#endregion

#region ==================== CHECK POINT (soft pause + trac) ====================

function Stop-CheckpointUI {
    if (-not $script:LastUIStop) { $script:LastUIStop = Get-Date "2000-01-01" }
    if ((New-TimeSpan -Start $script:LastUIStop -End (Get-Date)).TotalSeconds -lt $Config.UIStopMinIntervalSec) { return }

    Write-Log -Message ("UI-suppression: stoppar processer ({0})" -f ($Config.UIPatternsToStop -join ",")) -Level "INFO"
    foreach ($n in $Config.UIPatternsToStop) {
        Get-Process -Name $n -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    $script:LastUIStop = Get-Date
}

function Start-CheckpointServicesForce {
    Write-Log -Message ("Säkerställer att Check Point-tjänster kör: {0}" -f ($Config.CheckpointServicesStartOrder -join ", ")) -Level "INFO"

    foreach ($name in $Config.CheckpointServicesStartOrder) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if (-not $svc) {
            Write-Log -Message ("Tjänst saknas: {0}" -f $name) -Level "WARNING"
            continue
        }

        if ($svc.StartType -eq "Disabled") {
            Write-Log -Message ("Tjänst är Disabled, startar ej: {0}" -f $name) -Level "WARNING"
            continue
        }

        if ($svc.Status -ne "Running") {
            Write-Log -Message ("Start-Service: {0} ({1})" -f $svc.Name, $svc.DisplayName) -Level "INFO"
            try {
                Start-Service -Name $name -ErrorAction Stop
                $svc.WaitForStatus("Running", (New-TimeSpan -Seconds $Config.CheckpointServiceWaitSeconds))
                Write-Log -Message ("Startad: {0}" -f $name) -Level "SUCCESS"
            } catch {
                Write-Log -Message ("Start misslyckades: {0} | {1}" -f $name, $_.Exception.Message) -Level "WARNING"
            }
        } else {
            Write-Log -Message ("Redan igång: {0}" -f $name) -Level "INFO"
        }
    }
}

function Get-TracPath {
    $p = Join-Path -Path $Config.CheckPointPath -ChildPath $Config.TracExe
    if (Test-Path -Path $p -PathType Leaf) { return $p }
    return $null
}

function Invoke-Trac {
    param(
        [Parameter(Mandatory=$true)][string]$TracPath,
        [Parameter(Mandatory=$true)][string[]]$Args
    )
    $cmd = ('"{0}" {1}' -f $TracPath, ($Args -join " "))
    Write-Log -Message ("trac invoke: {0}" -f $cmd) -Level "INFO"
    try {
        $out = & $TracPath @Args 2>$null
        if ($out) {
            $lines = @($out | Select-Object -First 3)
            Write-Log -Message ("trac output (first lines): {0}" -f (($lines -join " | ").Trim())) -Level "INFO"
        } else {
            Write-Log -Message "trac output: <empty>" -Level "INFO"
        }
        return $out
    } catch {
        Write-Log -Message ("trac exception: {0}" -f $_.Exception.Message) -Level "WARNING"
        return $null
    }
}

function BestEffort-Disconnect {
    param([Parameter(Mandatory=$true)][string]$TracPath)
    Write-Log -Message "Soft pause: trac disconnect + UI suppression" -Level "INFO"
    $null = Invoke-Trac -TracPath $TracPath -Args @("disconnect")
    Stop-CheckpointUI
}

#endregion

#region ==================== CERT: selection via trac list ====================

function Get-TracCertificateListRaw {
    param([Parameter(Mandatory=$true)][string]$TracPath)

    $out = Invoke-Trac -TracPath $TracPath -Args @("list")
    if (-not $out) { return @() }

    @(
        $out |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -like "CN=*" }
    )
}

function Select-CertificatePath {
    param([Parameter(Mandatory=$true)][string[]]$TracList)

    $tag = $Config.FriendlyIssuerTag

    $candidates = $TracList | Where-Object { $_ -like "*(friendlyName=*($tag))*" }
    Write-Log -Message ("Cert candidates(tag={0}): {1}" -f $tag, $candidates.Count) -Level "INFO"

    $candidates = $candidates | Where-Object { $_ -notmatch $Config.AdminExcludeRegex }
    Write-Log -Message ("Cert candidates(after exclude): {0}" -f $candidates.Count) -Level "INFO"

    $preferred = $candidates | Where-Object { $_ -notmatch '(?i)\bCN=.*\s-\s' }
    if ($preferred.Count -ge 1) {
        $pick = $preferred | Select-Object -First 1
        Write-Log -Message ("Valde cert preferred: {0}" -f $pick) -Level "SUCCESS"
        return $pick
    }

    if ($candidates.Count -ge 1) {
        $pick = $candidates | Select-Object -First 1
        Write-Log -Message ("Valde cert fallback: {0}" -f $pick) -Level "WARNING"
        return $pick
    }

    $fallback = $TracList | Where-Object { $_ -notmatch $Config.AdminExcludeRegex }
    if ($fallback.Count -ge 1) {
        $pick = $fallback | Select-Object -First 1
        Write-Log -Message ("Valde cert last resort: {0}" -f $pick) -Level "WARNING"
        return $pick
    }

    Write-Log -Message "Kunde inte välja cert (ingen match)." -Level "ERROR"
    return $null
}

function Wait-ForSelectedCertificateViaTrac {
    param([Parameter(Mandatory=$true)][string]$TracPath)

    $start = Get-Date
    Write-Log -Message ("Väntar på rätt cert via trac list (max {0}s, poll {1}ms)..." -f $Config.CertWaitMaxSeconds, $Config.CertPollMs) -Level "INFO"

    while (((Get-Date) - $start).TotalSeconds -lt $Config.CertWaitMaxSeconds) {
        $list = Get-TracCertificateListRaw -TracPath $TracPath
        if ($list.Count -gt 0) {
            $cert = Select-CertificatePath -TracList $list
            if ($cert) {
                $elapsed = [int](((Get-Date) - $start).TotalSeconds)
                Write-Log -Message ("Rätt cert hittat efter {0}s." -f $elapsed) -Level "SUCCESS"
                return $cert
            }
        }
        Start-Sleep -Milliseconds $Config.CertPollMs
    }

    Write-Log -Message "Timeout: inget valbart cert hittades via trac." -Level "ERROR"
    return $null
}

#endregion

#region ==================== CONNECT (RECOMMENDED QUOTING) ====================

function Connect-CheckPointVPN {
    param(
        [Parameter(Mandatory=$true)][string]$TracPath,
        [Parameter(Mandatory=$true)][string]$CertificatePath
    )

    Write-Log -Message ("Initierar VPN connect: site={0}" -f $Config.VPNSite) -Level "INFO"
    Write-Log -Message ("certificate_path={0}" -f $CertificatePath) -Level "INFO"

    # Explicit quoting runt -s och -d (viktigt pga mellanslag och specialtecken i certsträngen)
    $siteEsc = $Config.VPNSite.Replace('"','\"')
    $certEsc = $CertificatePath.Replace('"','\"')
    $argLine = 'connect -s "{0}" -d "{1}"' -f $siteEsc, $certEsc

    Write-Log -Message ("Startar: `"{0}`" {1}" -f $TracPath, $argLine) -Level "INFO"

    $p = Start-Process -FilePath $TracPath -ArgumentList $argLine -PassThru -WindowStyle Hidden

    $finished = $true
    try {
        Wait-Process -Id $p.Id -Timeout $Config.VPNConnectTimeoutSec -ErrorAction Stop
    } catch {
        $finished = $false
    }

    if (-not $finished) {
        Write-Log -Message ("VPN connect timeout efter {0}s. Stoppar PID={1}" -f $Config.VPNConnectTimeoutSec, $p.Id) -Level "WARNING"
        try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch {}
        return $false
    }

    try {
        $p.Refresh()
        Write-Log -Message ("trac.exe exitcode={0}" -f $p.ExitCode) -Level "INFO"
        return ($p.ExitCode -eq 0)
    } catch {
        Write-Log -Message ("Kunde inte läsa exitcode: {0}" -f $_.Exception.Message) -Level "WARNING"
        return $false
    }
}

#endregion

#region ==================== MAIN ====================

$script:StartTime = Get-Date

try {
    Initialize-Logging

    Write-Log -Message "========================================" -Level "INFO"
    Write-Log -Message ("Start {0} | User={1} | Computer={2} | whoami={3} | PID={4}" -f $Config.LogPrefix, $env:USERNAME, $env:COMPUTERNAME, (whoami), $PID) -Level "INFO"
    Write-Log -Message "========================================" -Level "INFO"

    if (-not (Acquire-Lock)) { exit 0 }

    # 1) Starta Net iD
    Invoke-Step -Name "Start Net iD" -ScriptBlock { Start-NetId } | Out-Null

    # 2) Resolve trac.exe path
    $tracPath = Invoke-Step -Name "Resolve trac.exe path" -ScriptBlock {
        $tp = Get-TracPath
        if (-not $tp) { throw ("trac.exe saknas i: {0}" -f (Join-Path $Config.CheckPointPath $Config.TracExe)) }
        Write-Log -Message ("tracPath={0}" -f $tp) -Level "INFO"
        $tp
    }

    # 3) Soft pause (disconnect + UI-stop) och säkerställ tjänster (så trac list fungerar)
    Invoke-Step -Name "Soft pause (disconnect + UI) + ensure services" -ScriptBlock {
        Start-CheckpointServicesForce
        BestEffort-Disconnect -TracPath $tracPath
    } | Out-Null

    # 4) Inside-check (pre)
    $inside = Invoke-Step -Name "Inside-detection (pre)" -ScriptBlock { Test-InsideNetwork }
    if ($inside) {
        Write-Log -Message "Inside detected. VPN behövs ej. Avslutar." -Level "SUCCESS"
        exit 0
    }

    # 5) Välj rätt cert via trac list
    $certPath = Invoke-Step -Name "Select certificate via trac list" -ScriptBlock {
        $cert = Wait-ForSelectedCertificateViaTrac -TracPath $tracPath
        if (-not $cert) { throw "Kunde inte välja cert via trac." }
        $cert
    }

    # 6) Pre-connect UI suppression
    Invoke-Step -Name "Pre-connect UI suppression" -ScriptBlock { Stop-CheckpointUI } | Out-Null

    # 7) Connect med valt cert (citerad -d)
    $ok = Invoke-Step -Name "Connect VPN (cert controlled)" -ScriptBlock {
        Connect-CheckPointVPN -TracPath $tracPath -CertificatePath $certPath
    }

    if (-not $ok) {
        Write-Log -Message "Connect misslyckades." -Level "ERROR"
        exit 30
    }

    # 8) Post-connect wait + inside-check
    Invoke-Step -Name "Post-connect wait" -ScriptBlock {
        Write-Log -Message ("Väntar {0}s innan ny inside-check..." -f $Config.PostConnectWaitSeconds) -Level "INFO"
        Start-Sleep -Seconds $Config.PostConnectWaitSeconds
    } | Out-Null

    $insideAfter = Invoke-Step -Name "Inside-detection (post)" -ScriptBlock { Test-InsideNetwork }
    if ($insideAfter) {
        Write-Log -Message "Post-connect inside-check: OK (VPN uppe)." -Level "SUCCESS"
        exit 0
    } else {
        Write-Log -Message "Post-connect inside-check: FAIL (VPN ej verifierad via reachability)." -Level "WARNING"
        exit 31
    }
}
catch {
    Write-Log -Message ("Oväntat fel: {0}" -f $_.Exception.Message) -Level "ERROR"
    Write-Log -Message ("Stack: {0}" -f $_.ScriptStackTrace) -Level "ERROR"
    exit 99
}
finally {
    Release-Lock
    $elapsed = [int]((New-TimeSpan -Start $script:StartTime -End (Get-Date)).TotalSeconds)
    Write-Log -Message ("Slut. Total tid: {0}s" -f $elapsed) -Level "INFO"
}

#endregion
