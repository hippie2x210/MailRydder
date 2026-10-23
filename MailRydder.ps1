<#
  MailRydder - rydder Gmail-innboksen (og litt PC-søppel) for venner og familie.

  SIKKERHET
  - Bruker kun Gmail-tilgangene "gmail.modify" + "gmail.settings.basic".
    Med disse er det TEKNISK UMULIG å slette e-post permanent. Alt havner i Papirkurven
    og kan gjenopprettes i 30 dager. Scriptet kaller aldri delete/batchDelete og tømmer
    aldri papirkurven.
  - Stjernemerket e-post og e-post med ord som faktura/kvittering/forfall/purring osv. røres aldri.
  - Alt som flyttes logges til en CSV på skrivebordet.

  Kjør via Start-MailRydder.bat (eller: powershell -ExecutionPolicy Bypass -File .\MailRydder.ps1)
#>
[CmdletBinding()]
param(
    [switch]$SkipPC,          # hopp over PC-opprydding
    [switch]$SkipFilter,      # ikke lag Gmail-filter for fremtidig søppel
    [switch]$DryRun,          # vis bare hva som ville blitt gjort
    [int]$MaxScan = 2000,     # hvor mange av de nyeste innboks-e-postene som analyseres for "nyhetsbrev-avsendere"
    [int]$MinPerSender = 5    # en avsender med nyhetsbrev-header og minst så mange e-poster regnes som søppel
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptDir 'config.json'
$Desktop = [Environment]::GetFolderPath('Desktop')
$Stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$LogCsv = Join-Path $Desktop "MailRydder-logg_$Stamp.csv"
$ReportTxt = Join-Path $Desktop "MailRydder-rapport_$Stamp.txt"
$Report = New-Object System.Collections.Generic.List[string]
$Log = New-Object System.Collections.Generic.List[object]

function Say([string]$msg, [string]$color = 'Gray') { Write-Host $msg -ForegroundColor $color; $Report.Add($msg) }

# ----------------------------------------------------------------------------------------
# Regler
# ----------------------------------------------------------------------------------------
# Ord som gjør at en e-post ALDRI flyttes (lagt til som -{...} i alle søk)
$ProtectWords = @(
    'faktura','efaktura','invoice','kvittering','receipt','forfall','forfaller','purring','inkasso',
    'betalingspåminnelse','skatteetaten','skattemelding','nav','altinn','digipost','helsenorge',
    'kontrakt','signer','signere','e-signer','lønnsslipp','payslip','legetime','time hos',
    'billett','ticket','booking','bestillingsbekreftelse','ordrebekreftelse','order confirmation'
)
$Protect = '-is:starred -in:sent -{' + (($ProtectWords | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }) -join ' ') + '}'

# Faste søk som alltid er trygt søppel
$Rules = [ordered]@{
    'Reklame-fanen'                 = 'in:inbox category:promotions'
    'Sosialt-fanen (varsler)'       = 'in:inbox category:social'
    'Google "delt kontodata"-varsler' = 'in:inbox from:noreply-accounts@google.com'
    'Gamle engangskoder'            = 'in:inbox older_than:3d subject:(kode OR code OR "verification" OR verifiser OR sikkerhetskode OR påloggingskode OR "sign in" OR "sign-in" OR OTP)'
    'Velkommen-e-poster (eldre enn 14 dager)' = 'in:inbox older_than:14d subject:(velkommen OR welcome)'
    'Sosiale medier-mas'            = 'in:inbox from:(mail.instagram.com OR priority.instagram.com OR facebookmail.com OR email.snapchat.com OR tiktok.com OR linkedin.com OR pinterest.com OR twitter.com OR x.com) -subject:(innlogging OR login OR passord OR password OR security OR sikkerhet)'
    'Temu/AliExpress/Shein-mas'     = 'in:inbox from:(temuemail.com OR temu.com OR aliexpress.com OR shein.com OR wish.com) -subject:(henting OR "venter på henting" OR pickup)'
}

# ----------------------------------------------------------------------------------------
# Google-innlogging (OAuth for skrivebordsapper, med PKCE og localhost-tilbakekall)
# ----------------------------------------------------------------------------------------
$Scopes = 'https://www.googleapis.com/auth/gmail.modify https://www.googleapis.com/auth/gmail.settings.basic'

function Get-Config {
    $c = $null
    if (Test-Path $ConfigPath) {
        try { $c = Get-Content $ConfigPath -Raw | ConvertFrom-Json }
        catch { Say "Kunne ikke lese config.json. Veiviseren oppretter en ny fil." 'Yellow' }
    }
    if (-not $c -or -not $c.client_id -or $c.client_id -like '*PASTE*') {
        Say "`n=== Koble til Google første gang ===" 'Yellow'
        Say 'MailRydder trenger en OAuth-klient fra Google for å kunne vise Googles sikre innlogging.'
        Say 'Dette er en engangsoppsett for personen som deler programmet; Gmail-brukeren logger inn i nettleseren etterpå.'
        Say '1. Åpne Google Cloud Console og velg/opprett et prosjekt: https://console.cloud.google.com/'
        Say '2. Aktiver Gmail API, konfigurer samtykkeskjermen og legg Gmail-adressen til som testbruker.'
        Say '3. Opprett OAuth-klient av typen Desktop app under Google Auth Platform → Clients.'
        Say '4. Kopier klient-ID-en som slutter på .apps.googleusercontent.com.'
        Say 'Google-veiledning: https://developers.google.com/identity/protocols/oauth2/native-app' 'DarkCyan'
        try { Start-Process 'https://console.cloud.google.com/auth/clients' } catch {}
        $clientId = Read-Host 'Lim inn OAuth client ID her (Enter avbryter)'
        if (-not $clientId -or $clientId -notmatch '^[^\s]+\.apps\.googleusercontent\.com$') {
            throw 'Oppsett avbrutt. Klient-ID-en må slutte på .apps.googleusercontent.com.'
        }
        # Desktop-applikasjoner er offentlige OAuth-klienter. PKCE beskytter koden,
        # og Google dokumenterer client_secret som valgfri for denne flyten.
        $c = [pscustomobject]@{ client_id = $clientId; client_secret = '' }
        $c | ConvertTo-Json | Set-Content -Path $ConfigPath -Encoding UTF8
        Say "Lagret klient-ID i $ConfigPath" 'Green'
    }
    $secret = ''
    if ($c.PSObject.Properties['client_secret'] -and $c.client_secret -and $c.client_secret -notlike '*PASTE*') { $secret = $c.client_secret }
    return [pscustomobject]@{ client_id = $c.client_id; client_secret = $secret }
}

function New-Base64Url([byte[]]$bytes) { [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_') }

function Get-AccessToken($cfg) {
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    $buf = New-Object byte[] 32; $rng.GetBytes($buf)
    $verifier = New-Base64Url $buf
    $challenge = New-Base64Url ([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier)))
    $stateBuf = New-Object byte[] 16; $rng.GetBytes($stateBuf); $state = New-Base64Url $stateBuf

    # finn ledig port
    $tcp = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0); $tcp.Start()
    $port = $tcp.LocalEndpoint.Port; $tcp.Stop()
    $redirect = "http://localhost:$port/"

    $listener = New-Object Net.HttpListener
    $listener.Prefixes.Add($redirect)
    $listener.Start()

    $authUrl = 'https://accounts.google.com/o/oauth2/v2/auth?' + (@(
        "client_id=$([uri]::EscapeDataString($cfg.client_id))",
        "redirect_uri=$([uri]::EscapeDataString($redirect))",
        'response_type=code',
        "scope=$([uri]::EscapeDataString($Scopes))",
        "code_challenge=$challenge",
        'code_challenge_method=S256',
        "state=$state",
        'prompt=select_account consent'
    ) -join '&')

    Say "Åpner nettleseren for Google-innlogging... (logg inn med kontoen som skal ryddes)" 'Cyan'
    Start-Process $authUrl

    $ctx = $listener.GetContext()
    $q = $ctx.Request.QueryString
    $html = '<html><body style="font-family:sans-serif;text-align:center;padding-top:60px"><h2>Ferdig! Du kan lukke denne fanen og gå tilbake til MailRydder.</h2></body></html>'
    $bytes = [Text.Encoding]::UTF8.GetBytes($html)
    $ctx.Response.ContentType = 'text/html; charset=utf-8'
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length); $ctx.Response.Close()
    $listener.Stop()

    if ($q['error']) { throw "Google-innlogging avbrutt: $($q['error'])" }
    if ($q['state'] -ne $state) { throw 'Ugyldig svar fra Google (state stemmer ikke).' }

    $body = @{
        code = $q['code']; client_id = $cfg.client_id
        redirect_uri = $redirect; grant_type = 'authorization_code'; code_verifier = $verifier
    }
    if ($cfg.client_secret) { $body.client_secret = $cfg.client_secret }
    $tok = Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' -Body $body
    return $tok.access_token
}

# ----------------------------------------------------------------------------------------
# Gmail API-hjelpere (kun lese, merke, flytte til papirkurv, filtre)
# ----------------------------------------------------------------------------------------
$Api = 'https://gmail.googleapis.com/gmail/v1/users/me'
$script:Headers = $null

function Invoke-Gmail([string]$Method, [string]$Path, $Body = $null) {
    # Sperre: dette scriptet skal aldri kunne slette permanent.
    if ($Method -eq 'DELETE' -or $Path -match 'batchDelete' ) { throw 'Permanent sletting er sperret i MailRydder.' }
    $uri = if ($Path -like 'http*') { $Path } else { "$Api/$Path" }
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

function Find-MessageIds([string]$query, [int]$limit = 100000) {
    $ids = New-Object System.Collections.Generic.List[string]
    $page = $null
    do {
        $p = "messages?maxResults=500&q=$([uri]::EscapeDataString($query))"
        if ($page) { $p += "&pageToken=$page" }
        $r = Invoke-Gmail GET $p
        if ($r.PSObject.Properties['messages'] -and $r.messages) { foreach ($m in $r.messages) { $ids.Add($m.id) } }
        $page = if ($r.PSObject.Properties['nextPageToken']) { $r.nextPageToken } else { $null }
    } while ($page -and $ids.Count -lt $limit)
    return ,$ids
}

function Get-Meta([string]$id) {
    $m = Invoke-Gmail GET "messages/$id`?format=metadata&metadataHeaders=From&metadataHeaders=Subject&metadataHeaders=List-Unsubscribe&metadataHeaders=Date"
    $h = @{}
    if ($m.payload.PSObject.Properties['headers']) { foreach ($x in $m.payload.headers) { $h[$x.name.ToLower()] = $x.value } }
    $labels = if ($m.PSObject.Properties['labelIds']) { $m.labelIds } else { @() }
    [pscustomobject]@{
        Id = $id; From = $h['from']; Subject = $h['subject']; Date = $h['date']
        Unsub = [bool]$h['list-unsubscribe']; Starred = ($labels -contains 'STARRED')
    }
}

function Get-Email([string]$from) {
    if (-not $from) { return '' }
    if ($from -match '<([^>]+)>') { return $Matches[1].ToLower() }
    return $from.Trim().ToLower()
}

function Move-ToTrash([System.Collections.Generic.List[string]]$ids, [string]$reason) {
    if ($ids.Count -eq 0) { return 0 }
    foreach ($id in $ids) { $Log.Add([pscustomobject]@{ Regel = $reason; MessageId = $id; Lenke = "https://mail.google.com/mail/u/0/#all/$id" }) }
    if ($DryRun) { return $ids.Count }
    for ($i = 0; $i -lt $ids.Count; $i += 1000) {
        $chunk = $ids.GetRange($i, [math]::Min(1000, $ids.Count - $i))
        # Legger på TRASH-etiketten = "Flytt til papirkurv". Ingenting slettes permanent.
        try {
            Invoke-Gmail POST 'messages/batchModify' @{ ids = @($chunk); addLabelIds = @('TRASH'); removeLabelIds = @('INBOX') } | Out-Null
        } catch {
            # Reserve: en og en via "trash" (flytter til papirkurv, sletter ikke)
            foreach ($one in $chunk) { try { Invoke-Gmail POST "messages/$one/trash" | Out-Null } catch {} }
        }
    }
    return $ids.Count
}

function Get-OrCreateLabel([string]$name) {
    $labels = (Invoke-Gmail GET 'labels').labels
    $l = $labels | Where-Object { $_.name -eq $name } | Select-Object -First 1
    if ($l) { return $l.id }
    if ($DryRun) { return 'DRYRUN' }
    return (Invoke-Gmail POST 'labels' @{ name = $name; labelListVisibility = 'labelShow'; messageListVisibility = 'show' }).id
}

function New-JunkFilters([string[]]$senders, [string]$labelId) {
    # Fremtidig post fra disse avsenderne hopper over innboksen og får etiketten "Auto-sortert".
    # (Filtrene arkiverer bare - de sletter ingenting.)
    $existing = @()
    try { $f = Invoke-Gmail GET 'settings/filters'; if ($f.PSObject.Properties['filter']) { $existing = $f.filter } } catch {}
    $already = @{}
    foreach ($e in $existing) { if ($e.criteria.PSObject.Properties['from']) { foreach ($s in ($e.criteria.from -split ' OR ')) { $already[$s.Trim().ToLower()] = 1 } } }
    $todo = @($senders | Where-Object { -not $already.ContainsKey($_) } | Sort-Object -Unique)
    $made = 0; $chunk = @(); $len = 0
    foreach ($s in ($todo + '__END__')) {
        if ($s -eq '__END__' -or ($len + $s.Length) -gt 900) {
            if ($chunk.Count -gt 0) {
                if (-not $DryRun) {
                    Invoke-Gmail POST 'settings/filters' @{
                        criteria = @{ from = ($chunk -join ' OR ') }
                        action   = @{ addLabelIds = @($labelId); removeLabelIds = @('INBOX') }
                    } | Out-Null
                }
                $made++
            }
            $chunk = @(); $len = 0
            if ($s -eq '__END__') { break }
        }
        $chunk += $s; $len += $s.Length + 4
    }
    return @{ Filters = $made; Senders = $todo.Count }
}

# ----------------------------------------------------------------------------------------
# PC-opprydding (forsiktig: kun midlertidige filer. Papirkurven i Windows røres IKKE.)
# ----------------------------------------------------------------------------------------
function Invoke-PcCleanup {
    Say "`n=== PC-opprydding ===" 'Yellow'
    $drive = Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':'))
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
            try { Remove-Item $f.FullName -Force -ErrorAction Stop; $n++ } catch {}  # filer i bruk hoppes over
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
    Say '==============================================' 'Cyan'
    Say '  MailRydder  -  rydder innboksen trygt' 'Cyan'
    Say '  (alt havner i Papirkurven og kan angres i 30 dager)' 'Cyan'
    Say '==============================================' 'Cyan'
    if ($DryRun) { Say '  TESTMODUS: ingenting blir endret.' 'Magenta' }

    $cfg = Get-Config
    $token = Get-AccessToken $cfg
    $script:Headers = @{ Authorization = "Bearer $token" }
    $acct = Invoke-Gmail GET 'profile'
    Say ("`nInnlogget som {0}  ({1:N0} e-poster totalt)" -f $acct.emailAddress, $acct.messagesTotal) 'Green'

    $total = 0

    # 1) Faste regler
    Say "`n=== Faste regler ===" 'Yellow'
    foreach ($name in $Rules.Keys) {
        $ids = Find-MessageIds "$($Rules[$name]) $Protect"
        $n = Move-ToTrash $ids $name
        $total += $n
        Say ("  {0,-45} {1,6} flyttet til papirkurv" -f $name, $n)
    }

    # 2) "Like e-poster": avsendere med nyhetsbrev-header som sender mye
    Say "`n=== Finner nyhetsbrev-avsendere (analyserer opptil $MaxScan e-poster) ===" 'Yellow'
    $scanIds = Find-MessageIds "in:inbox $Protect" $MaxScan
    $metas = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($id in $scanIds) {
        $i++
        if ($i -gt $MaxScan) { break }
        if ($i % 25 -eq 0) { Write-Progress -Activity 'Analyserer innboksen' -Status "$i / $([math]::Min($MaxScan,$scanIds.Count))" -PercentComplete (100 * $i / [math]::Min($MaxScan, $scanIds.Count)) }
        try { $metas.Add((Get-Meta $id)) } catch {}
    }
    Write-Progress -Activity 'Analyserer innboksen' -Completed

    $junkSenders = @(
        $metas | Where-Object { $_.Unsub -and -not $_.Starred } |
        Group-Object { Get-Email $_.From } |
        Where-Object { $_.Name -and $_.Count -ge $MinPerSender } |
        Sort-Object Count -Descending
    )
    $senderList = New-Object System.Collections.Generic.List[string]
    foreach ($g in $junkSenders) {
        $ids = Find-MessageIds "in:inbox from:$($g.Name) $Protect"
        $n = Move-ToTrash $ids "Nyhetsbrev: $($g.Name)"
        $total += $n
        $senderList.Add($g.Name)
        Say ("  {0,-45} {1,6} flyttet" -f $g.Name, $n)
    }
    if ($junkSenders.Count -eq 0) { Say '  Ingen tydelige nyhetsbrev-avsendere funnet.' }

    # 3) Filter for fremtiden
    if (-not $SkipFilter -and $senderList.Count -gt 0) {
        Say "`n=== Lager filter for fremtidig søppel ===" 'Yellow'
        $labelId = Get-OrCreateLabel 'Auto-sortert'
        $res = New-JunkFilters $senderList.ToArray() $labelId
        Say ("  {0} nye avsendere sendes nå rett til etiketten 'Auto-sortert' (i {1} filter)." -f $res.Senders, $res.Filters) 'Green'
    }

    # 4) Viktig å sjekke
    Say "`n=== Bør sjekkes (ingenting av dette er flyttet) ===" 'Yellow'
    $checkQ = 'in:inbox newer_than:45d {faktura efaktura forfall forfaller purring inkasso "har du glemt" signer signere skatteetaten "venter på henting" "lagringsplassen"}'
    $checkIds = Find-MessageIds $checkQ 15
    if ($checkIds.Count -eq 0) { Say '  Ingenting funnet.' }
    foreach ($id in ($checkIds | Select-Object -First 15)) {
        try { $m = Get-Meta $id; Say ("  - {0}  |  {1}" -f ($m.From -replace '<.*>', '').Trim(), $m.Subject) 'White' } catch {}
    }

    Say ("`nFERDIG: {0:N0} e-poster flyttet til Papirkurven." -f $total) 'Green'
    Say '  Angre: åpne Gmail -> Papirkurv -> merk -> "Flytt til innboks". (Papirkurven tømmes IKKE av MailRydder.)'

    if (-not $SkipPC) { Invoke-PcCleanup }

    $Log | Export-Csv -Path $LogCsv -NoTypeInformation -Encoding UTF8
    $Report | Set-Content -Path $ReportTxt -Encoding UTF8
    Say "`nLogg:    $LogCsv" 'DarkGray'
    Say "Rapport: $ReportTxt" 'DarkGray'
}
catch {
    Write-Host "`nFEIL: $($_.Exception.Message)" -ForegroundColor Red
    try { $Log | Export-Csv -Path $LogCsv -NoTypeInformation -Encoding UTF8 } catch {}
}
finally {
    Write-Host "`nTrykk Enter for å lukke..." -ForegroundColor DarkGray
    [void](Read-Host)
}
