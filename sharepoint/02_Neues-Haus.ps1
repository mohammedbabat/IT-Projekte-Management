<#
.SYNOPSIS
    Legt ein neues Haus an und erzeugt automatisch dessen Standort-Checkliste aus der Vorlage.
    Kann außerdem die Fristen eines Hauses nach einer Terminverschiebung neu berechnen.

.DESCRIPTION
    Modus "Neu"         : Haus in "DHO Häuser" anlegen + alle aktiven Vorlagenpunkte nach "DHO Aufgaben" kopieren.
    Modus "Checkliste"  : Für ein bereits eingetragenes Haus (z. B. per Formular angelegt) die Checkliste erzeugen.
    Modus "Neuberechnen": Nach Änderung von Ü oder Go-Live alle offenen Fristen des Hauses neu berechnen.

    Frist = Übernahmetag (Ü) bzw. Go-Live (X) + Offset-Tage aus der Vorlage.
    Diese Logik ist identisch mit dem Power-Automate-Flow (siehe README) – das Skript ist die manuelle Variante / der Fallback.

.EXAMPLE
    # Neues Haus mit Checkliste
    .\02_Neues-Haus.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/IT-Projekte" -ClientId "<app-id>" `
        -Hotelname "Mercure Hotel Beispielstadt" -Ort "Beispielstadt" -Uebernahmetag 2027-02-01 -GoLive 2027-02-15

.EXAMPLE
    # Go-Live hat sich verschoben -> Datum im Haus ändern und Fristen neu berechnen
    .\02_Neues-Haus.ps1 -SiteUrl ... -ClientId ... -Ort "Beispielstadt" -GoLive 2027-03-01 -Neuberechnen

.EXAMPLE
    # Haus wurde in der Liste angelegt, aber ohne Checkliste (Flow aus) -> Checkliste nachziehen
    .\02_Neues-Haus.ps1 -SiteUrl ... -ClientId ... -Ort "Beispielstadt" -NurCheckliste
#>
[CmdletBinding(DefaultParameterSetName = 'Neu')]
param(
    [Parameter(Mandatory)] [string]$SiteUrl,
    [Parameter(Mandatory)] [string]$ClientId,
    [Parameter(Mandatory)] [string]$Ort,

    [Parameter(ParameterSetName = 'Neu', Mandatory)] [string]$Hotelname,
    [Parameter(ParameterSetName = 'Neu')] [string]$Hauskuerzel,
    [Parameter(ParameterSetName = 'Neu')] [string]$Direktor,
    [Parameter(ParameterSetName = 'Neu')] [string]$OraclePL,
    [Parameter(ParameterSetName = 'Neu')] [string]$AccorKontakt,

    [Parameter(ParameterSetName = 'Neu')]
    [Parameter(ParameterSetName = 'Neuberechnen')]
    [Nullable[datetime]]$Uebernahmetag,

    [Parameter(ParameterSetName = 'Neu')]
    [Parameter(ParameterSetName = 'Neuberechnen')]
    [Nullable[datetime]]$GoLive,

    [Parameter(ParameterSetName = 'Checkliste', Mandatory)] [switch]$NurCheckliste,
    [Parameter(ParameterSetName = 'Neuberechnen', Mandatory)] [switch]$Neuberechnen
)

$ErrorActionPreference = 'Stop'

$ListeHaeuser  = 'DHO Häuser'
$ListeVorlage  = 'DHO Checkliste Vorlage'
$ListeAufgaben = 'DHO Aufgaben'

# ------------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------------
function Get-LokalDatum($Wert) {
    if ($null -eq $Wert -or $Wert -eq '') { return $null }
    return ([datetime]$Wert).ToLocalTime().Date
}

function Get-Frist {
    param([string]$Bezug, $Offset, [Nullable[datetime]]$Ue, [Nullable[datetime]]$X)
    if ($null -eq $Offset -or $Offset -eq '') { return $null }
    switch ($Bezug) {
        'Ü' { if ($Ue) { return $Ue.AddDays([int]$Offset) } }
        'X' { if ($X)  { return $X.AddDays([int]$Offset) } }
    }
    return $null
}

function Get-Haus([string]$Ort) {
    $caml = "<View><Query><Where><Eq><FieldRef Name='Ort' /><Value Type='Text'>$([Security.SecurityElement]::Escape($Ort))</Value></Eq></Where></Query></View>"
    $treffer = @(Get-PnPListItem -List $ListeHaeuser -Query $caml)
    if ($treffer.Count -gt 1) { throw "Mehrere Häuser mit Ort '$Ort' gefunden – bitte eindeutig machen." }
    return $treffer | Select-Object -First 1
}

function Get-AufgabenVonHaus([int]$HausID) {
    $caml = "<View><Query><Where><Eq><FieldRef Name='HausID' /><Value Type='Number'>$HausID</Value></Eq></Where></Query></View>"
    return @(Get-PnPListItem -List $ListeAufgaben -Query $caml -PageSize 500)
}

function New-Checkliste {
    param($Haus)
    $hausId = $Haus.Id
    $ort    = $Haus['Ort']
    $ue     = Get-LokalDatum $Haus['Uebernahmetag']
    $x      = Get-LokalDatum $Haus['GoLive']

    $vorhanden = Get-AufgabenVonHaus $hausId
    if ($vorhanden.Count -gt 0) {
        throw "Für '$ort' existieren bereits $($vorhanden.Count) Aufgaben. Abbruch, damit keine doppelte Checkliste entsteht."
    }

    $vorlage = @(Get-PnPListItem -List $ListeVorlage -PageSize 500 | Where-Object { $_['Aktiv'] -ne $false } |
        Sort-Object { [int]$_['Nr'] })
    if ($vorlage.Count -eq 0) { throw "Vorlage '$ListeVorlage' ist leer. Erst 01_DHO-Setup.ps1 ausführen." }

    $batch = New-PnPBatch
    $ohneFrist = 0
    foreach ($v in $vorlage) {
        $werte = @{
            Title      = $v['Title']
            Haus       = $ort
            HausID     = $hausId
            Nr         = [int]$v['Nr']
            Bereich    = $v['Bereich']
            Owner      = $v['Owner']
            Bezug      = $v['Bezug']
            OffsetTage = $v['OffsetTage']
            FristText  = $v['FristText']
            Hinweis    = $v['Hinweis']
            Status     = 'offen'
        }
        $frist = Get-Frist -Bezug $v['Bezug'] -Offset $v['OffsetTage'] -Ue $ue -X $x
        if ($frist) { $werte.Frist = $frist }
        elseif ($v['Bezug'] -ne 'fix') { $ohneFrist++ }
        Add-PnPListItem -List $ListeAufgaben -Values $werte -Batch $batch | Out-Null
    }
    Invoke-PnPBatch -Batch $batch
    Set-PnPListItem -List $ListeHaeuser -Identity $hausId -Values @{ ChecklisteErstellt = $true } | Out-Null

    Write-Host "Checkliste für '$ort' erzeugt: $($vorlage.Count) Punkte." -ForegroundColor Green
    if ($ohneFrist -gt 0) {
        Write-Host "$ohneFrist Punkte ohne Frist, weil Ü oder Go-Live noch fehlt. Nach dem Eintragen mit -Neuberechnen nachziehen." -ForegroundColor Yellow
    }
}

function Update-Fristen {
    param($Haus)
    $ue = Get-LokalDatum $Haus['Uebernahmetag']
    $x  = Get-LokalDatum $Haus['GoLive']
    $aufgaben = Get-AufgabenVonHaus $Haus.Id | Where-Object { $_['Status'] -notin @('erledigt', 'entfällt') }

    $batch = New-PnPBatch
    $geaendert = 0
    foreach ($a in $aufgaben) {
        $neu = Get-Frist -Bezug $a['Bezug'] -Offset $a['OffsetTage'] -Ue $ue -X $x
        $alt = Get-LokalDatum $a['Frist']
        if ($neu -ne $alt) {
            Set-PnPListItem -List $ListeAufgaben -Identity $a.Id -Values @{ Frist = $neu } -Batch $batch | Out-Null
            $geaendert++
        }
    }
    if ($geaendert -gt 0) { Invoke-PnPBatch -Batch $batch }
    Write-Host "Fristen für '$($Haus['Ort'])' neu berechnet: $geaendert von $($aufgaben.Count) offenen Punkten geändert." -ForegroundColor Green
}

# ------------------------------------------------------------------
# Ablauf
# ------------------------------------------------------------------
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

switch ($PSCmdlet.ParameterSetName) {
    'Neu' {
        if (Get-Haus $Ort) { throw "Haus mit Ort '$Ort' existiert bereits. Für Datumsänderungen -Neuberechnen verwenden." }
        $werte = @{
            Title              = $Hotelname
            Ort                = $Ort
            Hauskuerzel        = $Hauskuerzel
            Direktor           = $Direktor
            OraclePL           = $OraclePL
            AccorKontakt       = $AccorKontakt
            ChecklisteErstellt = $true   # verhindert, dass der Flow parallel eine zweite Checkliste erzeugt
        }
        if ($Uebernahmetag) { $werte.Uebernahmetag = $Uebernahmetag.Date }
        if ($GoLive)        { $werte.GoLive = $GoLive.Date }
        $haus = Add-PnPListItem -List $ListeHaeuser -Values $werte
        Write-Host "Haus '$Ort' angelegt (ID $($haus.Id))." -ForegroundColor Green
        $haus = Get-PnPListItem -List $ListeHaeuser -Id $haus.Id
        try { New-Checkliste $haus }
        catch {
            Set-PnPListItem -List $ListeHaeuser -Identity $haus.Id -Values @{ ChecklisteErstellt = $false } | Out-Null
            throw
        }
    }
    'Checkliste' {
        $haus = Get-Haus $Ort
        if (-not $haus) { throw "Haus '$Ort' nicht gefunden." }
        New-Checkliste $haus
    }
    'Neuberechnen' {
        $haus = Get-Haus $Ort
        if (-not $haus) { throw "Haus '$Ort' nicht gefunden." }
        $werte = @{}
        if ($Uebernahmetag) { $werte.Uebernahmetag = $Uebernahmetag.Date }
        if ($GoLive)        { $werte.GoLive = $GoLive.Date }
        if ($werte.Count -gt 0) {
            Set-PnPListItem -List $ListeHaeuser -Identity $haus.Id -Values $werte | Out-Null
            Write-Host "Termine für '$Ort' aktualisiert." -ForegroundColor Green
            $haus = Get-PnPListItem -List $ListeHaeuser -Id $haus.Id
        }
        Update-Fristen $haus
    }
}
