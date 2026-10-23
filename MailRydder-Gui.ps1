$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:workerGmail    = Join-Path $scriptDir 'MailRydder.ps1'
$script:workerExchange = Join-Path $scriptDir 'MailRydder-Exchange.ps1'

$form = New-Object System.Windows.Forms.Form
$form.Text            = 'MailRydder'
$form.StartPosition   = 'CenterScreen'
$form.ClientSize      = New-Object System.Drawing.Size(650, 560)
$form.MinimumSize     = New-Object System.Drawing.Size(666, 599)
$form.BackColor       = [System.Drawing.Color]::FromArgb(247, 249, 252)
$form.Font            = New-Object System.Drawing.Font('Segoe UI', 10)

# Tittel
$title = New-Object System.Windows.Forms.Label
$title.Text      = 'MailRydder'
$title.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 24)
$title.ForeColor = [System.Drawing.Color]::FromArgb(27, 48, 76)
$title.Location  = New-Object System.Drawing.Point(34, 24)
$title.Size      = New-Object System.Drawing.Size(570, 45)
$form.Controls.Add($title)

$intro = New-Object System.Windows.Forms.Label
$intro.Text     = 'Gå gjennom innboksen i ditt tempo. Start gjerne med en forhåndsvisning.'
$intro.Location = New-Object System.Drawing.Point(38, 77)
$intro.Size     = New-Object System.Drawing.Size(570, 24)
$form.Controls.Add($intro)

# E-postkonto-valg
$providerBox = New-Object System.Windows.Forms.GroupBox
$providerBox.Text     = 'E-postkonto'
$providerBox.Location = New-Object System.Drawing.Point(34, 112)
$providerBox.Size     = New-Object System.Drawing.Size(578, 64)
$form.Controls.Add($providerBox)

$script:rdoGmail = New-Object System.Windows.Forms.RadioButton
$script:rdoGmail.Text     = 'Gmail  (Google-konto)'
$script:rdoGmail.Location = New-Object System.Drawing.Point(16, 26)
$script:rdoGmail.Size     = New-Object System.Drawing.Size(250, 24)
$script:rdoGmail.Checked  = $true
$providerBox.Controls.Add($script:rdoGmail)

$script:rdoExchange = New-Object System.Windows.Forms.RadioButton
$script:rdoExchange.Text     = 'Exchange / Outlook  (Microsoft 365, Outlook.com)'
$script:rdoExchange.Location = New-Object System.Drawing.Point(280, 26)
$script:rdoExchange.Size     = New-Object System.Drawing.Size(280, 24)
$providerBox.Controls.Add($script:rdoExchange)

# Informasjonstekst
$info = New-Object System.Windows.Forms.Label
$info.Text      = "Forhåndsvisning endrer ingenting. Når du velger Rydd, flyttes treff til papirkurven (Gmail) eller Slettede elementer (Exchange). De kan gjenopprettes i opptil 30 dager."
$info.Location  = New-Object System.Drawing.Point(38, 192)
$info.Size      = New-Object System.Drawing.Size(570, 54)
$info.ForeColor = [System.Drawing.Color]::FromArgb(65, 78, 96)
$form.Controls.Add($info)

# Avkrysningsbokser (Gmail-spesifikk er skjult ved Exchange-valg)
$script:filter = New-Object System.Windows.Forms.CheckBox
$script:filter.Text     = 'Lag Gmail-filter for nyhetsbrev som blir funnet (arkiverer framtidige e-poster)'
$script:filter.Location = New-Object System.Drawing.Point(40, 260)
$script:filter.Size     = New-Object System.Drawing.Size(570, 36)
$script:filter.Checked  = $false
$form.Controls.Add($script:filter)

$script:pc = New-Object System.Windows.Forms.CheckBox
$script:pc.Text     = 'Rydd midlertidige filer på denne PC-en som er eldre enn 2 dager'
$script:pc.Location = New-Object System.Drawing.Point(40, 301)
$script:pc.Size     = New-Object System.Drawing.Size(570, 36)
$script:pc.Checked  = $false
$form.Controls.Add($script:pc)

# Skjul/vis Gmail-filter-valget basert på leverandørvalg
$updateFilterVisibility = {
    $script:filter.Visible = $script:rdoGmail.Checked
}
$script:rdoGmail.Add_CheckedChanged($updateFilterVisibility)
$script:rdoExchange.Add_CheckedChanged($updateFilterVisibility)

# Knapper
$preview = New-Object System.Windows.Forms.Button
$preview.Text      = 'Forhåndsvisning'
$preview.Location  = New-Object System.Drawing.Point(40, 370)
$preview.Size      = New-Object System.Drawing.Size(270, 54)
$preview.FlatStyle = 'Flat'
$preview.BackColor = [System.Drawing.Color]::FromArgb(225, 235, 247)
$preview.Add_Click({
    $worker = if ($script:rdoGmail.Checked) { $script:workerGmail } else { $script:workerExchange }
    $args   = @('-DryRun', '-SkipPC', '-SkipFilter')
    $form.Enabled = $false
    try {
        Start-Process powershell.exe -ArgumentList (@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $worker + '"')) + $args) -WorkingDirectory $scriptDir -Wait
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Kunne ikke starte forhåndsvisningen.`n`n$($_.Exception.Message)",'MailRydder',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally { $form.Enabled = $true; $form.Activate() }
})
$form.Controls.Add($preview)

$run = New-Object System.Windows.Forms.Button
$run.Text      = 'Rydd innboks'
$run.Location  = New-Object System.Drawing.Point(330, 370)
$run.Size      = New-Object System.Drawing.Size(270, 54)
$run.FlatStyle = 'Flat'
$run.BackColor = [System.Drawing.Color]::FromArgb(38, 103, 171)
$run.ForeColor = [System.Drawing.Color]::White
$run.Add_Click({
    $provider = if ($script:rdoGmail.Checked) { 'Gmail' } else { 'Exchange/Outlook' }
    $message  = "MailRydder vil flytte treff til papirkurven ($provider). Du kan angre dette innen 30 dager."
    if ($script:rdoGmail.Checked -and $script:filter.Checked) { $message += "`n`nDet opprettes Gmail-filter for nyhetsbrev som blir funnet." }
    if ($script:pc.Checked) { $message += "`n`nMidlertidige PC-filer eldre enn 2 dager blir slettet." }
    $answer = [System.Windows.Forms.MessageBox]::Show($message,"Bekreft opprydding",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($answer -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $worker = if ($script:rdoGmail.Checked) { $script:workerGmail } else { $script:workerExchange }
    $args   = @()
    if ($script:rdoGmail.Checked -and -not $script:filter.Checked) { $args += '-SkipFilter' }
    if (-not $script:pc.Checked) { $args += '-SkipPC' }

    $form.Enabled = $false
    try {
        Start-Process powershell.exe -ArgumentList (@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $worker + '"')) + $args) -WorkingDirectory $scriptDir -Wait
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Kunne ikke starte MailRydder.`n`n$($_.Exception.Message)",'MailRydder',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally { $form.Enabled = $true; $form.Activate() }
})
$form.Controls.Add($run)

$footer = New-Object System.Windows.Forms.Label
$footer.Text      = 'Nettleseren åpnes for innlogging. Resultatet vises i et eget vindu; trykk Enter der når du er ferdig. MailRydder ber aldri om passordet ditt.'
$footer.Location  = New-Object System.Drawing.Point(40, 448)
$footer.Size      = New-Object System.Drawing.Size(570, 42)
$footer.ForeColor = [System.Drawing.Color]::FromArgb(65, 78, 96)
$form.Controls.Add($footer)

[void]$form.ShowDialog()
