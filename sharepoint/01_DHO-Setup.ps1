<#
.SYNOPSIS
    Legt die SharePoint-Listen für das DHO-Übernahme-Projektmanagement an und importiert
    die Checklisten-Vorlage sowie (optional) den aktuellen Stand der bestehenden Häuser.

.DESCRIPTION
    Erstellt drei Listen:
      1. "DHO Häuser"            – ein Eintrag pro Hotelübernahme (Ü-Tag, Go-Live, Ansprechpartner)
      2. "DHO Checkliste Vorlage" – die Master-Vorlage (82 Punkte, Bezug Ü/X + Offset-Tage)
      3. "DHO Aufgaben"           – die Standort-Checklisten aller Häuser (eine Zeile pro Punkt und Haus)

    Neue Häuser bekommen ihre Checkliste danach automatisch (Power Automate, siehe README)
    oder per 02_Neues-Haus.ps1.

.PARAMETER SiteUrl
    SharePoint-Site, auf der die Listen angelegt werden.
.PARAMETER ClientId
    Client-ID der Entra-App-Registrierung für PnP.PowerShell (Pflicht seit PnP 2.x).
.PARAMETER CsvPfad
    Ordner mit DHO_Vorlage.csv, DHO_Haeuser.csv, DHO_Aufgaben_Bestand.csv.
.PARAMETER BestandImportieren
    Importiert die bestehenden Häuser
    inkl. aktuellem Status aus der Excel-Datei "IT-Checklisten_Standorte".

.EXAMPLE
    .\01_DHO-Setup.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/IT-Projekte" -ClientId "xxxxxxxx-..." -BestandImportieren
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$SiteUrl,
    [Parameter(Mandatory)] [string]$ClientId,
    [string]$CsvPfad,
    [switch]$BestandImportieren
)

$ErrorActionPreference = 'Stop'

# CSV-Ordner bestimmen: Parameter -> Ordner des Skripts -> aktueller Ordner
if ([string]::IsNullOrWhiteSpace($CsvPfad)) { $CsvPfad = $PSScriptRoot }
if ([string]::IsNullOrWhiteSpace($CsvPfad)) { $CsvPfad = (Get-Location).Path }
if (-not (Test-Path (Join-Path $CsvPfad 'DHO_Vorlage.csv'))) {
    throw "DHO_Vorlage.csv nicht gefunden in '$CsvPfad'. Bitte -CsvPfad angeben, z. B. -CsvPfad 'C:\DHO'."
}
Write-Host "CSV-Ordner: $CsvPfad" -ForegroundColor DarkGray

# ------------------------------------------------------------------
# Konfiguration
# ------------------------------------------------------------------
$ListeHaeuser  = 'DHO Häuser'
$ListeVorlage  = 'DHO Checkliste Vorlage'
$ListeAufgaben = 'DHO Aufgaben'
$StatusWerte   = @('offen', 'prüfen', 'klären', 'erledigt', 'entfällt')

# ------------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------------
function Write-Schritt([string]$Text) { Write-Host "`n==> $Text" -ForegroundColor Cyan }

function New-DHOListe {
    param([string]$Titel, [string]$Url, [string]$TitelAnzeigename)
    $liste = Get-PnPList -Identity $Titel -ErrorAction SilentlyContinue
    if ($liste) {
        Write-Host "   Liste '$Titel' existiert bereits – wird weiterverwendet." -ForegroundColor Yellow
    }
    else {
        $liste = New-PnPList -Title $Titel -Url $Url -Template GenericList -OnQuickLaunch
        Write-Host "   Liste '$Titel' angelegt." -ForegroundColor Green
    }
    if ($TitelAnzeigename) {
        Set-PnPField -List $Titel -Identity 'Title' -Values @{ Title = $TitelAnzeigename } | Out-Null
    }
    return $liste
}

function Add-DHOFeld {
    param([string]$Liste, [string]$InternalName, [string]$FieldXml)
    $vorhanden = Get-PnPField -List $Liste -Identity $InternalName -ErrorAction SilentlyContinue
    if ($vorhanden) { return }
    Add-PnPFieldFromXml -List $Liste -FieldXml $FieldXml | Out-Null
    Write-Host "   + Feld $InternalName" -ForegroundColor DarkGray
}

function ConvertTo-ChoiceXml([string[]]$Werte) {
    ($Werte | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ''
}

function ConvertTo-Datum([string]$Wert) {
    if ([string]::IsNullOrWhiteSpace($Wert)) { return $null }
    return [datetime]::ParseExact($Wert, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}

function Set-DHOStandardansicht {
    param([string]$Liste, [string[]]$Felder)
    $ansicht = Get-PnPView -List $Liste | Where-Object DefaultView
    Set-PnPView -List $Liste -Identity $ansicht.Id -Fields $Felder | Out-Null
}

function Add-DHOAnsicht {
    param([string]$Liste, [string]$Titel, [string[]]$Felder, [string]$Query, [int]$RowLimit = 500)
    if (Get-PnPView -List $Liste -Identity $Titel -ErrorAction SilentlyContinue) {
        Remove-PnPView -List $Liste -Identity $Titel -Force
    }
    Add-PnPView -List $Liste -Title $Titel -Fields $Felder -Query $Query -RowLimit $RowLimit -Paged | Out-Null
    Write-Host "   + Ansicht '$Titel'" -ForegroundColor DarkGray
}

# ------------------------------------------------------------------
# Vorprüfung
# ------------------------------------------------------------------
foreach ($datei in 'DHO_Vorlage.csv', 'DHO_Haeuser.csv', 'DHO_Aufgaben_Bestand.csv') {
    if (-not (Test-Path (Join-Path $CsvPfad $datei))) {
        if ($datei -eq 'DHO_Vorlage.csv' -or $BestandImportieren) { throw "Datei fehlt: $(Join-Path $CsvPfad $datei)" }
    }
}
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'Modul PnP.PowerShell fehlt. Installation: Install-Module PnP.PowerShell -Scope CurrentUser'
}

Write-Schritt "Verbinde mit $SiteUrl"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# ------------------------------------------------------------------
# 1. Liste "DHO Häuser"
# ------------------------------------------------------------------
Write-Schritt "Liste '$ListeHaeuser'"
New-DHOListe -Titel $ListeHaeuser -Url 'Lists/DHOHaeuser' -TitelAnzeigename 'Hotelname' | Out-Null

Add-DHOFeld $ListeHaeuser 'Ort'              '<Field Type="Text" Name="Ort" StaticName="Ort" DisplayName="Ort" Required="TRUE" />'
Add-DHOFeld $ListeHaeuser 'Hauskuerzel'      '<Field Type="Text" Name="Hauskuerzel" StaticName="Hauskuerzel" DisplayName="Hauskürzel (z. B. DHO501)" />'
Add-DHOFeld $ListeHaeuser 'Uebernahmetag'    '<Field Type="DateTime" Name="Uebernahmetag" StaticName="Uebernahmetag" DisplayName="Übernahmetag (Ü)" Format="DateOnly" />'
Add-DHOFeld $ListeHaeuser 'GoLive'           '<Field Type="DateTime" Name="GoLive" StaticName="GoLive" DisplayName="Go-Live Hotelsystem (X)" Format="DateOnly" />'
Add-DHOFeld $ListeHaeuser 'Direktor'         '<Field Type="Text" Name="Direktor" StaticName="Direktor" DisplayName="Direktor/in" />'
Add-DHOFeld $ListeHaeuser 'PMSChampion'      '<Field Type="Text" Name="PMSChampion" StaticName="PMSChampion" DisplayName="PMS Champion" />'
Add-DHOFeld $ListeHaeuser 'OraclePL'         '<Field Type="Text" Name="OraclePL" StaticName="OraclePL" DisplayName="Oracle-Projektleitung" />'
Add-DHOFeld $ListeHaeuser 'AccorKontakt'     '<Field Type="Note" Name="AccorKontakt" StaticName="AccorKontakt" DisplayName="Accor-Ansprechpartner" NumLines="3" RichText="FALSE" />'
Add-DHOFeld $ListeHaeuser 'Projektstatus'    "<Field Type=`"Choice`" Name=`"Projektstatus`" StaticName=`"Projektstatus`" DisplayName=`"Projektstatus`" Format=`"Dropdown`"><Default>in Vorbereitung</Default><CHOICES>$(ConvertTo-ChoiceXml @('in Vorbereitung','übernommen','live','abgeschlossen'))</CHOICES></Field>"
Add-DHOFeld $ListeHaeuser 'ChecklisteErstellt' '<Field Type="Boolean" Name="ChecklisteErstellt" StaticName="ChecklisteErstellt" DisplayName="Checkliste erstellt"><Default>0</Default></Field>'

Set-DHOStandardansicht $ListeHaeuser @('LinkTitle', 'Ort', 'Hauskuerzel', 'Uebernahmetag', 'GoLive', 'Direktor', 'OraclePL', 'Projektstatus', 'ChecklisteErstellt')

# ------------------------------------------------------------------
# 2. Liste "DHO Checkliste Vorlage"
# ------------------------------------------------------------------
Write-Schritt "Liste '$ListeVorlage'"
New-DHOListe -Titel $ListeVorlage -Url 'Lists/DHOChecklisteVorlage' -TitelAnzeigename 'Aufgabe' | Out-Null

Add-DHOFeld $ListeVorlage 'Nr'         '<Field Type="Number" Name="Nr" StaticName="Nr" DisplayName="Nr." Decimals="0" Required="TRUE" />'
Add-DHOFeld $ListeVorlage 'Bereich'    '<Field Type="Text" Name="Bereich" StaticName="Bereich" DisplayName="Bereich" />'
Add-DHOFeld $ListeVorlage 'Owner'      '<Field Type="Text" Name="Owner" StaticName="Owner" DisplayName="Wen frage ich?" />'
Add-DHOFeld $ListeVorlage 'Bezug'      "<Field Type=`"Choice`" Name=`"Bezug`" StaticName=`"Bezug`" DisplayName=`"Bezug (Ü/X)`" Format=`"Dropdown`"><Default>X</Default><CHOICES>$(ConvertTo-ChoiceXml @('Ü','X','fix'))</CHOICES></Field>"
Add-DHOFeld $ListeVorlage 'OffsetTage' '<Field Type="Number" Name="OffsetTage" StaticName="OffsetTage" DisplayName="Offset Tage" Decimals="0" />'
Add-DHOFeld $ListeVorlage 'FristText'  '<Field Type="Text" Name="FristText" StaticName="FristText" DisplayName="Frist-Text (sofort/laufend)" />'
Add-DHOFeld $ListeVorlage 'Hinweis'    '<Field Type="Note" Name="Hinweis" StaticName="Hinweis" DisplayName="Hinweis" NumLines="3" RichText="FALSE" />'
Add-DHOFeld $ListeVorlage 'Aktiv'      '<Field Type="Boolean" Name="Aktiv" StaticName="Aktiv" DisplayName="Aktiv"><Default>1</Default></Field>'

Set-DHOStandardansicht $ListeVorlage @('Nr', 'Bereich', 'LinkTitle', 'Owner', 'Bezug', 'OffsetTage', 'FristText', 'Hinweis', 'Aktiv')
Add-DHOAnsicht -Liste $ListeVorlage -Titel 'Nach Bereich' `
    -Felder @('Nr', 'LinkTitle', 'Owner', 'Bezug', 'OffsetTage', 'FristText', 'Hinweis', 'Aktiv') `
    -Query '<GroupBy Collapse="FALSE"><FieldRef Name="Bereich" /></GroupBy><OrderBy><FieldRef Name="Nr" /></OrderBy>'

# ------------------------------------------------------------------
# 3. Liste "DHO Aufgaben"
# ------------------------------------------------------------------
Write-Schritt "Liste '$ListeAufgaben'"
New-DHOListe -Titel $ListeAufgaben -Url 'Lists/DHOAufgaben' -TitelAnzeigename 'Aufgabe' | Out-Null

Add-DHOFeld $ListeAufgaben 'Haus'       '<Field Type="Text" Name="Haus" StaticName="Haus" DisplayName="Haus" Required="TRUE" Indexed="TRUE" />'
Add-DHOFeld $ListeAufgaben 'HausID'     '<Field Type="Number" Name="HausID" StaticName="HausID" DisplayName="Haus-ID" Decimals="0" Indexed="TRUE" />'
Add-DHOFeld $ListeAufgaben 'Nr'         '<Field Type="Number" Name="Nr" StaticName="Nr" DisplayName="Nr." Decimals="0" />'
Add-DHOFeld $ListeAufgaben 'Bereich'    '<Field Type="Text" Name="Bereich" StaticName="Bereich" DisplayName="Bereich" />'
Add-DHOFeld $ListeAufgaben 'Owner'      '<Field Type="Text" Name="Owner" StaticName="Owner" DisplayName="Wen frage ich?" />'
Add-DHOFeld $ListeAufgaben 'Bezug'      "<Field Type=`"Choice`" Name=`"Bezug`" StaticName=`"Bezug`" DisplayName=`"Bezug (Ü/X)`" Format=`"Dropdown`"><CHOICES>$(ConvertTo-ChoiceXml @('Ü','X','fix'))</CHOICES></Field>"
Add-DHOFeld $ListeAufgaben 'OffsetTage' '<Field Type="Number" Name="OffsetTage" StaticName="OffsetTage" DisplayName="Offset Tage" Decimals="0" />'
Add-DHOFeld $ListeAufgaben 'Frist'      '<Field Type="DateTime" Name="Frist" StaticName="Frist" DisplayName="Bis wann?" Format="DateOnly" Indexed="TRUE" />'
Add-DHOFeld $ListeAufgaben 'FristText'  '<Field Type="Text" Name="FristText" StaticName="FristText" DisplayName="Frist-Text" />'
Add-DHOFeld $ListeAufgaben 'Status'     "<Field Type=`"Choice`" Name=`"Status`" StaticName=`"Status`" DisplayName=`"Status`" Format=`"Dropdown`" Indexed=`"TRUE`"><Default>offen</Default><CHOICES>$(ConvertTo-ChoiceXml $StatusWerte)</CHOICES></Field>"
Add-DHOFeld $ListeAufgaben 'ErledigtAm' '<Field Type="DateTime" Name="ErledigtAm" StaticName="ErledigtAm" DisplayName="Erledigt am" Format="DateOnly" />'
Add-DHOFeld $ListeAufgaben 'Notiz'      '<Field Type="Note" Name="Notiz" StaticName="Notiz" DisplayName="Antwort / Notiz" NumLines="4" RichText="FALSE" AppendOnly="FALSE" />'
Add-DHOFeld $ListeAufgaben 'Hinweis'    '<Field Type="Note" Name="Hinweis" StaticName="Hinweis" DisplayName="Hinweis" NumLines="3" RichText="FALSE" />'

$felderAufgaben = @('Haus', 'Nr', 'LinkTitle', 'Owner', 'Frist', 'FristText', 'Status', 'ErledigtAm', 'Notiz', 'Hinweis')
Set-DHOStandardansicht $ListeAufgaben $felderAufgaben

$offenFilter = '<And><Neq><FieldRef Name="Status" /><Value Type="Choice">erledigt</Value></Neq><Neq><FieldRef Name="Status" /><Value Type="Choice">entfällt</Value></Neq></And>'

Add-DHOAnsicht -Liste $ListeAufgaben -Titel 'Nach Haus' -Felder $felderAufgaben `
    -Query '<GroupBy Collapse="TRUE"><FieldRef Name="Haus" /><FieldRef Name="Bereich" /></GroupBy><OrderBy><FieldRef Name="Nr" /></OrderBy>'

Add-DHOAnsicht -Liste $ListeAufgaben -Titel 'Überfällig' -Felder $felderAufgaben `
    -Query "<Where><And><Lt><FieldRef Name=`"Frist`" /><Value Type=`"DateTime`"><Today /></Value></Lt>$offenFilter</And></Where><GroupBy Collapse=`"FALSE`"><FieldRef Name=`"Haus`" /></GroupBy><OrderBy><FieldRef Name=`"Frist`" /></OrderBy>"

Add-DHOAnsicht -Liste $ListeAufgaben -Titel 'Nächste 14 Tage' -Felder $felderAufgaben `
    -Query "<Where><And><And><Geq><FieldRef Name=`"Frist`" /><Value Type=`"DateTime`"><Today /></Value></Geq><Leq><FieldRef Name=`"Frist`" /><Value Type=`"DateTime`"><Today OffsetDays=`"14`" /></Value></Leq></And>$offenFilter</And></Where><OrderBy><FieldRef Name=`"Frist`" /></OrderBy>"

Add-DHOAnsicht -Liste $ListeAufgaben -Titel 'Meine Aufgaben (ich)' -Felder $felderAufgaben `
    -Query "<Where><And><Contains><FieldRef Name=`"Owner`" /><Value Type=`"Text`">ich</Value></Contains>$offenFilter</And></Where><GroupBy Collapse=`"FALSE`"><FieldRef Name=`"Haus`" /></GroupBy><OrderBy><FieldRef Name=`"Frist`" /></OrderBy>"

Add-DHOAnsicht -Liste $ListeAufgaben -Titel 'Prüfen / Klären' -Felder $felderAufgaben `
    -Query '<Where><Or><Eq><FieldRef Name="Status" /><Value Type="Choice">prüfen</Value></Eq><Eq><FieldRef Name="Status" /><Value Type="Choice">klären</Value></Eq></Or></Where><GroupBy Collapse="FALSE"><FieldRef Name="Haus" /></GroupBy><OrderBy><FieldRef Name="Nr" /></OrderBy>'

# Farbliche Hervorhebung der Spalte "Bis wann?" (rot = überfällig, orange = <= 7 Tage)
$fristFormat = @'
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "txtContent": "@currentField",
  "style": {
    "padding": "2px 6px",
    "border-radius": "4px",
    "background-color": "=if(@currentField == '', '', if([$Status] == 'erledigt' || [$Status] == 'entfällt', '', if(@currentField < @now, '#fde7e9', if(@currentField <= @now + 604800000, '#fff4ce', ''))))",
    "color": "=if(@currentField == '', '', if([$Status] == 'erledigt' || [$Status] == 'entfällt', '', if(@currentField < @now, '#a80000', '')))"
  }
}
'@
Set-PnPField -List $ListeAufgaben -Identity 'Frist' -Values @{ CustomFormatter = $fristFormat } | Out-Null

$statusFormat = @'
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "txtContent": "@currentField",
  "style": {
    "padding": "2px 8px",
    "border-radius": "10px",
    "display": "inline-block",
    "background-color": "=if(@currentField == 'erledigt', '#dff6dd', if(@currentField == 'entfällt', '#f3f2f1', if(@currentField == 'prüfen' || @currentField == 'klären', '#fff4ce', '#eff6fc')))"
  }
}
'@
Set-PnPField -List $ListeAufgaben -Identity 'Status' -Values @{ CustomFormatter = $statusFormat } | Out-Null

# ------------------------------------------------------------------
# 4. Vorlage importieren
# ------------------------------------------------------------------
Write-Schritt 'Importiere Checklisten-Vorlage'
$vorhandeneVorlage = @(Get-PnPListItem -List $ListeVorlage -PageSize 500)
if ($vorhandeneVorlage.Count -gt 0) {
    Write-Host "   Vorlage enthält bereits $($vorhandeneVorlage.Count) Einträge – Import übersprungen." -ForegroundColor Yellow
}
else {
    $vorlage = Import-Csv -Path (Join-Path $CsvPfad 'DHO_Vorlage.csv') -Delimiter ';' -Encoding UTF8
    $batch = New-PnPBatch
    foreach ($z in $vorlage) {
        $werte = @{
            Title     = $z.Aufgabe
            Nr        = [int]$z.Nr
            Bereich   = $z.Bereich
            Owner     = $z.Owner
            Bezug     = $z.Bezug
            FristText = $z.FristText
            Hinweis   = $z.Hinweis
            Aktiv     = $true
        }
        if ($z.OffsetTage -ne '') { $werte.OffsetTage = [int]$z.OffsetTage }
        Add-PnPListItem -List $ListeVorlage -Values $werte -Batch $batch | Out-Null
    }
    Invoke-PnPBatch -Batch $batch
    Write-Host "   $($vorlage.Count) Vorlagenpunkte importiert." -ForegroundColor Green
}

# ------------------------------------------------------------------
# 5. Bestand importieren (optional)
# ------------------------------------------------------------------
if ($BestandImportieren) {
    Write-Schritt 'Importiere bestehende Häuser und deren Checklisten'
    $haeuser  = Import-Csv -Path (Join-Path $CsvPfad 'DHO_Haeuser.csv') -Delimiter ';' -Encoding UTF8
    $aufgaben = Import-Csv -Path (Join-Path $CsvPfad 'DHO_Aufgaben_Bestand.csv') -Delimiter ';' -Encoding UTF8
    $bestehend = @(Get-PnPListItem -List $ListeHaeuser -PageSize 500)

    foreach ($h in $haeuser) {
        if ($bestehend | Where-Object { $_['Ort'] -eq $h.Ort }) {
            Write-Host "   $($h.Ort): existiert bereits – übersprungen." -ForegroundColor Yellow
            continue
        }

        # ChecklisteErstellt = true, damit der Power-Automate-Flow keine zweite Checkliste erzeugt
        $hWerte = @{
            Title              = $h.Hotelname
            Ort                = $h.Ort
            Direktor           = $h.Direktor
            PMSChampion        = $h.PMSChampion
            OraclePL           = $h.OraclePL
            AccorKontakt       = $h.AccorKontakt
            ChecklisteErstellt = $true
        }
        $ue = ConvertTo-Datum $h.Uebernahmetag
        $x  = ConvertTo-Datum $h.GoLive
        if ($ue) { $hWerte.Uebernahmetag = $ue }
        if ($x)  { $hWerte.GoLive = $x }
        $haus = Add-PnPListItem -List $ListeHaeuser -Values $hWerte

        $batch = New-PnPBatch
        $anzahl = 0
        foreach ($a in ($aufgaben | Where-Object Ort -eq $h.Ort)) {
            $status = if ($StatusWerte -contains $a.Status) { $a.Status } else { 'offen' }
            $aWerte = @{
                Title     = $a.Aufgabe
                Haus      = $h.Ort
                HausID    = $haus.Id
                Nr        = [int]$a.Nr
                Bereich   = $a.Bereich
                Owner     = $a.Owner
                Bezug     = $a.Bezug
                FristText = $a.FristText
                Status    = $status
                Notiz     = $a.Notiz
                Hinweis   = $a.Hinweis
            }
            if ($a.OffsetTage -ne '') { $aWerte.OffsetTage = [int]$a.OffsetTage }
            $frist = ConvertTo-Datum $a.Frist
            if ($frist) { $aWerte.Frist = $frist }
            Add-PnPListItem -List $ListeAufgaben -Values $aWerte -Batch $batch | Out-Null
            $anzahl++
        }
        Invoke-PnPBatch -Batch $batch
        Write-Host "   $($h.Ort): Haus angelegt (ID $($haus.Id)), $anzahl Aufgaben importiert." -ForegroundColor Green
    }
}

Write-Schritt 'Fertig'
Write-Host "Listen: $SiteUrl/Lists/DHOHaeuser | $SiteUrl/Lists/DHOChecklisteVorlage | $SiteUrl/Lists/DHOAufgaben" -ForegroundColor Green
Write-Host 'Nächster Schritt: Power-Automate-Flow "Neues Haus -> Checkliste" einrichten (README) oder 02_Neues-Haus.ps1 nutzen.'
