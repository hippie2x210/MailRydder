<#
  MailRydder-Exchange - rydder Exchange/Outlook-innboksen trygt.
  Støtter Microsoft 365, Exchange Online og Outlook.com via Microsoft Graph API.

  SIKKERHET
  - Bruker kun Mail.ReadWrite-tilgang (kan ikke slette permanent).
  - Meldinger flyttes til "Slettede elementer", ikke permanent slettet.
  - Stjernemerket/flagget e-post og meldinger med beskyttelsesord røres aldri.
  - Alt som flyttes logges til en CSV på skrivebordet.

  Kjør via Start-MailRydder.bat eller:
    powershell -ExecutionPolicy Bypass -File .\MailRydder-Exchange.ps1
#>
[CmdletBinding()]
param(
    [switch]$SkipPC,
    [switch]$DryRun,
    [int]$MaxScan = 2000,
    [int]$MinPerSender = 5
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptDir 'config.json'
$Desktop    = [Environment]::GetFolderPath('Desktop')
$Stamp      = Get-Date -Format 'yyyy-MM-dd_HHmm'
$LogCsv     = Join-Path $Desktop "MailRydder-Exchange-logg_$Stamp.csv"
$ReportTxt  = Join-Path $Desktop "MailRydder-Exchange-rapport_$Stamp.txt"
$Report     = New-Object System.Collections.Generic.List[string]
$Log        = New-Object System.Collections.Generic.List[object]

function Say([string]$msg, [string]$color = 'Gray') { Write-Host $msg -ForegroundColor $color; $Report.Add($msg) }

# ----------------------------------------------------------------------------------------
# Beskyttelse - disse meldingene flyttes ALDRI
# ----------------------------------------------------------------------------------------
$ProtectWords = @(
    'faktura','efaktura','invoice','kvittering','receipt','forfall','forfaller','purring','inkasso',
    'betalingspåminnelse','skatteetaten','skattemelding','nav','altinn','digipost','helsenorge',
    'kontrakt','signer','signere','e-signer','lønnsslipp','payslip','legetime','time hos',
    'billett','ticket','booking','bestillingsbekreftelse','ordrebekreftelse','order confirmation'
)

$JunkDomains = @(
    'mail.instagram.com','priority.instagram.com','facebookmail.com','email.snapchat.com',
    'tiktok.com','linkedin.com','pinterest.com','twitter.com','x.com',
    'temuemail.com','temu.com','aliexpress.com','shein.com','wish.com'
)

$SafeSubjectWords = @('innlogging','login','passord','password','security','sikkerhet')

# ----------------------------------------------------------------------------------------
# Konfigurasjon og OAuth (PKCE, Microsoft identity platform)
# ----------------------------------------------------------------------------------------
function Get-ExchangeConfig {
    $c = $null
    if (Test-Path $ConfigPath) {
        try { $c = Get-Content $ConfigPath -Raw | ConvertFrom-Json } catch {}
    }

    $xCfg = $null
    if ($c -and $c.PSObject.Properties['exchange']) { $xCfg = $c.exchange }

    if (-not $xCfg -or -not $xCfg.client_id) {
        Say "`n=== Koble til Microsoft første gang ===" 'Yellow'
        Say 'MailRydder trenger en Azure-app-registrering for å koble til Exchange/Outlook.'
        Say ''
        Say '1. Gå til Azure-portalen:  https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade'
        Say '2. Klikk "New registration".'
        Say '   - Navn: MailRydder (eller valgfritt)'
        Say '   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"'
        Say '3. Under "Redirect URI": velg "Public client/native (mobile & desktop)", skriv inn: http://localhost'
        Say '4. Gå til "API permissions" → Add a permission → Microsoft Graph → Delegated:'
        Say '   Legg til: Mail.ReadWrite'
        Say '5. Kopier "Application (client) ID" fra oversiktssiden.'
        Say ''
        try { Start-Process 'https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade' } catch {}
        $clientId = Read-Host 'Lim inn Application (client) ID her (Enter avbryter)'
        if (-not $clientId -or $clientId -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            throw 'Ugyldig klient-ID. Formatet skal være: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
        }

        if (-not $c) { $c = [pscustomobject]@{} }
        $newXCfg = [pscustomobject]@{ client_id = $clientId; tenant_id = 'common' }
        if ($c.PSObject.Properties['exchange']) {
            $c.exchange = $newXCfg
        } else {
            $c | Add-Member -NotePropertyName 'exchange' -NotePropertyValue $newXCfg
        }
        $c | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding UTF8
        Say "Lagret klient-ID i $ConfigPath" 'Green'
        $xCfg = $newXCfg
    }
    return $xCfg
}

function New-Base64Url([byte[]]$bytes) { [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_') }

function Get-ExchangeToken($cfg) {
    $tenantId = if ($cfg.PSObject.Properties['tenant_id'] -and $cfg.tenant_id) { $cfg.tenant_id } else { 'common' }

    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    $buf = New-Object byte[] 32; $rng.GetBytes($buf)
    $verifier   = New-Base64Url $buf
    $challenge  = New-Base64Url ([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier)))
    $stateBuf   = New-Object byte[] 16; $rng.GetBytes($stateBuf)
    $state      = New-Base64Url $stateBuf

    $tcp  = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0); $tcp.Start()
    $port = $tcp.LocalEndpoint.Port; $tcp.Stop()
    $redirect = "http://localhost:$port/"

    $listener = New-Object Net.HttpListener
    $listener.Prefixes.Add($redirect)
    $listener.Start()

    $scopes  = 'Mail.ReadWrite offline_access'
    $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize?" + (@(
        "client_id=$([uri]::EscapeDataString($cfg.client_id))",
        "redirect_uri=$([uri]::EscapeDataString($redirect))",
        'response_type=code',
        "scope=$([uri]::EscapeDataString($scopes))",
        "code_challenge=$challenge",
        'code_challenge_method=S256',
        "state=$state",
        'prompt=select_account'
    ) -join '&')

    Say 'Åpner nettleseren for Microsoft-innlogging...' 'Cyan'
    Start-Process $authUrl

    $ctx   = $listener.GetContext()
    $q     = $ctx.Request.QueryString
    $html  = '<html><body style="font-family:sans-serif;text-align:center;padding-top:60px"><h2>Ferdig! Du kan lukke denne fanen og gå tilbake til MailRydder.</h2></body></html>'
    $bytes = [Text.Encoding]::UTF8.GetBytes($html)
    $ctx.Response.ContentType = 'text/html; charset=utf-8'
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length); $ctx.Response.Close()
    $listener.Stop()

    if ($q['error'])        { throw "Microsoft-innlogging avbrutt: $($q['error']) - $($q['error_description'])" }
    if ($q['state'] -ne $state) { throw 'Ugyldig svar fra Microsoft (state stemmer ikke).' }

    $body = @{
        code          = $q['code']
        client_id     = $cfg.client_id
        redirect_uri  = $redirect
        grant_type    = 'authorization_code'
        code_verifier = $verifier
        scope         = $scopes
    }
    $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body -ContentType 'application/x-www-form-urlencoded'
    return $tok.access_token
}

# ----------------------------------------------------------------------------------------
# Microsoft Graph API-hjelpere
# ----------------------------------------------------------------------------------------
$GraphBase   = 'https://graph.microsoft.com/v1.0'
$script:Headers = $null

function Invoke-Graph([string]$Method, [string]$Path, $Body = $null) {
    $uri = if ($Path -like 'https://*') { $Path } else { "$GraphBase/$Path" }
    for ($try = 1; $try -le 5; $try++) {
        try {
            if ($Body) {
                $json = $Body | ConvertTo-Json -Depth 10 -Compress
                return Invoke-RestMethod -Method $Method -Uri $uri -Headers $script:Headers -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($json))
            } else {
                return Invoke-RestMethod -Method $Method -Uri $uri -Headers $script:Headers
            }
        } catch {
            $code = $null
            try { $code = [int]$_.Exception.Response.StatusCode } catch {}
            if ($code -in 429, 500, 502, 503 -and $try -lt 5) { Start-Sleep -Seconds ([math]::Pow(2, $try)); continue }
            throw
        }
    }
}

function Get-InboxMessages([int]$limit = 100000) {
    $messages = New-Object System.Collections.Generic.List[object]
    $select = 'id,from,subject,receivedDateTime,internetMessageHeaders,flag'
    $uri = "me/mailFolders/inbox/messages?`$top=100&`$select=$select&`$orderby=receivedDateTime desc"

    $prog = 0
    do {
        $r = Invoke-Graph GET $uri
        if ($r.value) {
            foreach ($m in $r.value) { $messages.Add($m) }
        }
        $prog++
        if ($prog % 5 -eq 0) {
            Write-Progress -Activity 'Laster innboksen' -Status "$($messages.Count) meldinger hentet" -PercentComplete ([math]::Min(99, 100 * $messages.Count / $limit))
        }
        $uri = if ($r.'@odata.nextLink') { $r.'@odata.nextLink' } else { $null }
    } while ($uri -and $messages.Count -lt $limit)

    Write-Progress -Activity 'Laster innboksen' -Completed
    return ,$messages
}

function Test-Protected($msg) {
    if ($msg.flag -and $msg.flag.flagStatus -eq 'flagged') { return $true }
    $subj = if ($msg.subject) { $msg.subject.ToLower() } else { '' }
    foreach ($w in $ProtectWords) { if ($subj.Contains($w.ToLower())) { return $true } }
    return $false
}

function Has-UnsubscribeHeader($msg) {
    if (-not $msg.internetMessageHeaders) { return $false }
    foreach ($h in $msg.internetMessageHeaders) {
        if ($h.name -ieq 'List-Unsubscribe' -and $h.value) { return $true }
    }
    return $false
}

function Get-SenderAddress($msg) {
    if ($msg.from -and $msg.from.emailAddress -and $msg.from.emailAddress.address) {
        return $msg.from.emailAddress.address.ToLower()
    }
    return ''
}

function Move-ToDeletedItems([System.Collections.Generic.List[object]]$messages, [string]$reason) {
    if ($messages.Count -eq 0) { return 0 }
    foreach ($m in $messages) {
        $Log.Add([pscustomobject]@{
            Regel = $reason
            MessageId = $m.id
            Fra = (Get-SenderAddress $m)
            Emne = $m.subject
            Dato = $m.receivedDateTime
        })
    }
    if ($DryRun) { return $messages.Count }

    # Batch-API: maks 20 operasjoner per kall
    $moved = 0
    for ($i = 0; $i -lt $messages.Count; $i += 20) {
        $chunk = $messages.GetRange($i, [math]::Min(20, $messages.Count - $i))
        $reqs  = @()
        $idx   = 1
        foreach ($m in $chunk) {
            $reqs += [pscustomobject]@{
                id      = "$idx"
                method  = 'POST'
                url     = "/me/messages/$($m.id)/move"
                headers = [pscustomobject]@{ 'Content-Type' = 'application/json' }
                body    = [pscustomobject]@{ destinationId = 'deleteditems' }
            }
            $idx++
        }
        try {
            $batchRes = Invoke-RestMethod -Method POST `
                -Uri "$GraphBase/`$batch" `
                -Headers $script:Headers `
                -ContentType 'application/json; charset=utf-8' `
                -Body ([Text.Encoding]::UTF8.GetBytes((@{ requests = $reqs } | ConvertTo-Json -Depth 10 -Compress)))
            $moved += ($batchRes.responses | Where-Object { $_.status -in 200, 201 }).Count
        } catch {
            # Reserve: en og en
            foreach ($m in $chunk) {
                try { Invoke-Graph POST "me/messages/$($m.id)/move" @{ destinationId = 'deleteditems' } | Out-Null; $moved++ } catch {}
            }
        }
    }
    return $moved
}

# ----------------------------------------------------------------------------------------
# PC-opprydding (identisk med Gmail-versjonen)
# ----------------------------------------------------------------------------------------
function Invoke-PcCleanup {
    Say "`n=== PC-opprydding ===" 'Yellow'
    $drive  = Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':'))
    $before = $drive.Free
    $targets = @(
        $env:TEMP,
        (Join-Path $env:LOCALAPPDATA 'Temp'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\INetCache'),
        (Join-Path $env:LOCALAPPDATA 'CrashDumps'),
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data\Default\Cache'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data\Default\Cache')
    ) | Where-Object { $_ -and (Test-Path $_) } | Sort-Object -Unique
    foreach ($t in $targets) {
        $files = Get-ChildItem $t -Recurse -Force -File -ErrorAction SilentlyContinue |
                 Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-2) }
        $n = 0
        foreach ($f in $files) {
            if ($DryRun) { $n++; continue }
            try { Remove-Item $f.FullName -Force -ErrorAction Stop; $n++ } catch {}
        }
        Say ("  {0,-70} {1} filer fjernet" -f $t, $n)
    }
    $after = (Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':'))).Free
    Say ("  Frigjort plass: {0:N0} MB  (ledig nå: {1:N1} GB)" -f (($after - $before) / 1MB), ($after / 1GB)) 'Green'
    Say "`n  Programmer som starter automatisk (bare info - ingenting endret):" 'Yellow'
    try {
        Get-CimInstance Win32_StartupCommand -ErrorAction Stop |
            Select-Object -First 25 | ForEach-Object { Say ("   - {0}" -f $_.Name) }
    } catch { Say '   (kunne ikke lese oppstartsprogrammer)' }
}

# ----------------------------------------------------------------------------------------
# Hovedprogram
# ----------------------------------------------------------------------------------------
try {
    Clear-Host
    Say '===============================================' 'Cyan'
    Say '  MailRydder Exchange  -  rydder innboksen trygt' 'Cyan'
    Say '  (flyttes til Slettede elementer, kan angres)' 'Cyan'
    Say '===============================================' 'Cyan'
    if ($DryRun) { Say '  TESTMODUS: ingenting blir endret.' 'Magenta' }

    $cfg    = Get-ExchangeConfig
    $token  = Get-ExchangeToken $cfg
    $script:Headers = @{ Authorization = "Bearer $token" }

    $me    = Invoke-Graph GET 'me?$select=displayName,mail,userPrincipalName'
    $email = if ($me.mail) { $me.mail } else { $me.userPrincipalName }
    Say ("`nInnlogget som {0}  ({1})" -f $me.displayName, $email) 'Green'

    # Last alle innboks-meldinger én gang (brukes av alle regler)
    Say "`nLaster innboksen (maks $MaxScan meldinger)..." 'Gray'
    $allMsgs = Get-InboxMessages $MaxScan
    Say ("  {0} meldinger lastet." -f $allMsgs.Count)

    $total = 0

    # 1) Faste regler
    Say "`n=== Faste regler ===" 'Yellow'

    # Gamle engangskoder (>3 dager)
    $otpKw   = @('kode','code','verification','verifiser','sikkerhetskode','påloggingskode','sign in','otp','one-time','einmalig')
    $cutOtp  = (Get-Date).AddDays(-3)
    $otpHits = New-Object System.Collections.Generic.List[object]
    foreach ($m in $allMsgs) {
        if ([datetime]$m.receivedDateTime -gt $cutOtp) { continue }
        if (Test-Protected $m) { continue }
        $subj = if ($m.subject) { $m.subject.ToLower() } else { '' }
        foreach ($kw in $otpKw) { if ($subj.Contains($kw)) { $otpHits.Add($m); break } }
    }
    $n = Move-ToDeletedItems $otpHits 'Gamle engangskoder (>3 dager)'; $total += $n
    Say ("  {0,-47} {1,6} flyttet" -f 'Gamle engangskoder (>3 dager)', $n)

    # Velkomst-e-poster (>14 dager)
    $welKw   = @('velkommen','welcome','bienvenue','willkommen')
    $cutWel  = (Get-Date).AddDays(-14)
    $welHits = New-Object System.Collections.Generic.List[object]
    foreach ($m in $allMsgs) {
        if ([datetime]$m.receivedDateTime -gt $cutWel) { continue }
        if (Test-Protected $m) { continue }
        $subj = if ($m.subject) { $m.subject.ToLower() } else { '' }
        foreach ($kw in $welKw) { if ($subj.Contains($kw)) { $welHits.Add($m); break } }
    }
    $n = Move-ToDeletedItems $welHits 'Velkomst-e-poster (>14 dager)'; $total += $n
    Say ("  {0,-47} {1,6} flyttet" -f 'Velkomst-e-poster (>14 dager)', $n)

    # Sosiale medier-mas (kjente domener, ikke sikkerhetsrelatert)
    $socialHits = New-Object System.Collections.Generic.List[object]
    foreach ($m in $allMsgs) {
        if (Test-Protected $m) { continue }
        $from = Get-SenderAddress $m
        $subj = if ($m.subject) { $m.subject.ToLower() } else { '' }
        $junk = $false
        foreach ($d in $JunkDomains) { if ($from.EndsWith("@$d") -or $from.EndsWith(".$d")) { $junk = $true; break } }
        if (-not $junk) { continue }
        $safe = $false
        foreach ($kw in $SafeSubjectWords) { if ($subj.Contains($kw)) { $safe = $true; break } }
        if (-not $safe) { $socialHits.Add($m) }
    }
    $n = Move-ToDeletedItems $socialHits 'Sosiale medier-mas'; $total += $n
    Say ("  {0,-47} {1,6} flyttet" -f 'Sosiale medier-mas', $n)

    # 2) Nyhetsbrev-avsendere (List-Unsubscribe-header + minst $MinPerSender meldinger)
    Say "`n=== Finner nyhetsbrev-avsendere (analyserer $($allMsgs.Count) meldinger) ===" 'Yellow'

    $junkSenders = @(
        $allMsgs |
        Where-Object { -not (Test-Protected $_) -and (Has-UnsubscribeHeader $_) } |
        Group-Object { Get-SenderAddress $_ } |
        Where-Object { $_.Name -and $_.Count -ge $MinPerSender } |
        Sort-Object Count -Descending
    )

    foreach ($g in $junkSenders) {
        $hits = New-Object System.Collections.Generic.List[object]
        foreach ($m in $allMsgs) {
            if (Test-Protected $m) { continue }
            if ((Get-SenderAddress $m) -eq $g.Name) { $hits.Add($m) }
        }
        $n = Move-ToDeletedItems $hits "Nyhetsbrev: $($g.Name)"; $total += $n
        Say ("  {0,-47} {1,6} flyttet" -f $g.Name, $n)
    }
    if ($junkSenders.Count -eq 0) { Say '  Ingen tydelige nyhetsbrev-avsendere funnet.' }

    # 3) Viktig å sjekke (ikke rørt)
    Say "`n=== Bør sjekkes (ingenting av dette er flyttet) ===" 'Yellow'
    $checkKw  = @('faktura','forfall','purring','inkasso','signer','skatteetaten','venter på henting','lagringsplassen')
    $cutCheck = (Get-Date).AddDays(-45)
    $shown    = 0
    foreach ($m in $allMsgs) {
        if ([datetime]$m.receivedDateTime -lt $cutCheck) { continue }
        $subj = if ($m.subject) { $m.subject.ToLower() } else { '' }
        $match = $false
        foreach ($kw in $checkKw) { if ($subj.Contains($kw.ToLower())) { $match = $true; break } }
        if (-not $match) { continue }
        $fromName = if ($m.from -and $m.from.emailAddress) { $m.from.emailAddress.name } else { '' }
        Say ("  - {0}  |  {1}" -f $fromName, $m.subject) 'White'
        $shown++
        if ($shown -ge 15) { break }
    }
    if ($shown -eq 0) { Say '  Ingenting funnet.' }

    Say ("`nFERDIG: {0:N0} e-poster flyttet til Slettede elementer." -f $total) 'Green'
    Say '  Angre: åpne Outlook → Slettede elementer → merk → "Flytt til innboksen".'

    if (-not $SkipPC) { Invoke-PcCleanup }

    $Log    | Export-Csv -Path $LogCsv    -NoTypeInformation -Encoding UTF8
    $Report | Set-Content -Path $ReportTxt -Encoding UTF8
    Say "`nLogg:    $LogCsv"    'DarkGray'
    Say "Rapport: $ReportTxt" 'DarkGray'
} catch {
    Write-Host "`nFEIL: $($_.Exception.Message)" -ForegroundColor Red
    try { $Log | Export-Csv -Path $LogCsv -NoTypeInformation -Encoding UTF8 } catch {}
} finally {
    Write-Host "`nTrykk Enter for å lukke..." -ForegroundColor DarkGray
    [void](Read-Host)
}
