# IT-Projekte-Management – DHO Hotelübernahmen

Dashboard und Automatisierung für die IT-Checklisten bei Hotelübernahmen.

> **Wichtig:** Dieses Repository ist öffentlich. Es enthält **keine Projektdaten**
> (keine Excel-/CSV-Dateien, keine Namen, keine Termine). Die Daten bleiben in
> SharePoint bzw. lokal; das Dashboard liest sie nur im Browser.
> `.gitignore` blockiert `*.xlsx`, `*.xlsm`, `*.csv`.

## 1. Dashboard (`index.html`)

Aufruf über GitHub Pages. Zwei Datenquellen: **live aus SharePoint** (siehe Abschnitt 4) oder Excel-/CSV-Datei:

1. `IT-Checklisten_Standorte.xlsx` per Drag & Drop laden (ein Blatt pro Haus + Blatt „Vorlage“)
2. optional zusätzlich `PM_Checkliste_DHO_Uebernahmen.xlsx` für „Nächste Termine“
3. alternativ: CSV-Export `DHO_Haeuser.csv` + `DHO_Aufgaben_Bestand.csv`

Inhalt:

- Kennzahlen: Fortschritt, überfällig, fällig in 7 Tagen, nächster Go-Live
- Fortschritt je Haus (erledigt / prüfen-klären / überfällig / offen)
- Zeitplan Übernahmetag (Ü) und Go-Live (X) mit Heute-Linie
- Heatmap „% erledigt je Bereich und Haus“
- Aufgabenliste mit Filtern (Haus, überfällig, 14 Tage, Verantwortliche, Suche)
- **Generator „Neues Haus → Checkliste“**: erzeugt aus der Vorlage eine komplette
  Standort-Checkliste mit berechneten Fristen (Excel-Blatt oder CSV für SharePoint)

Die Daten werden **nur im Browser** gelesen und lokal (localStorage) gemerkt.
Über „Daten vergessen“ werden sie wieder entfernt.

## 2. SharePoint-Listen (`sharepoint/`)

| Skript | Zweck |
|---|---|
| `01_DHO-Setup.ps1` | legt die Listen `DHO Häuser`, `DHO Checkliste Vorlage`, `DHO Aufgaben` inkl. Ansichten an, importiert Vorlage und optional den Bestand |
| `02_Neues-Haus.ps1` | neues Haus + Checkliste anlegen, Checkliste nachziehen oder Fristen nach Terminverschiebung neu berechnen |

Voraussetzungen: `Install-Module PnP.PowerShell -Scope CurrentUser` und eine
Entra-App-Registrierung für PnP (Client-ID).

```powershell
.\01_DHO-Setup.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>" -ClientId "<app-id>" -CsvPfad "C:\DHO\Daten" -BestandImportieren
.\02_Neues-Haus.ps1 -SiteUrl ... -ClientId ... -Hotelname "Hotel X" -Ort "Stadt" -Uebernahmetag 2027-02-01 -GoLive 2027-02-15
.\02_Neues-Haus.ps1 -SiteUrl ... -ClientId ... -Ort "Stadt" -GoLive 2027-03-01 -Neuberechnen
```

Die CSV-Dateien für den Import liegen **nicht** im Repo, sondern lokal.

**Fristlogik** (wie in der Excel-Vorlage): `Frist = Übernahmetag (Ü) bzw. Go-Live (X) + Offset-Tage`.
Punkte mit Bezug `fix` haben einen Text („sofort“, „laufend“).

## 3. Power Automate – Checkliste automatisch bei neuem Haus

### Flow A: „DHO – Neues Haus / Termin geändert“

1. **Trigger:** SharePoint – *Wenn ein Element erstellt oder geändert wird*, Liste `DHO Häuser`
2. **Bedingung:** `ChecklisteErstellt` ist gleich `false`

   **Ja-Zweig (Checkliste erzeugen):**
   1. *Elemente abrufen* – Liste `DHO Checkliste Vorlage`, Filter `Aktiv eq 1`, Top 500
   2. *Auf alle anwenden* → *Element erstellen* in `DHO Aufgaben`:
      - Titel = `Title`, Haus = `Ort` (Trigger), HausID = `ID` (Trigger)
      - Nr, Bereich, Owner, Bezug, OffsetTage, FristText, Hinweis aus der Vorlage
      - Status = `offen`
      - Bis wann? (Frist) = Ausdruck:
        ```
        if(equals(items('Auf_alle_anwenden')?['Bezug']?['Value'],'Ü'),
           if(or(empty(triggerBody()?['Uebernahmetag']),empty(items('Auf_alle_anwenden')?['OffsetTage'])),null,
              addDays(triggerBody()?['Uebernahmetag'],int(items('Auf_alle_anwenden')?['OffsetTage']),'yyyy-MM-dd')),
           if(equals(items('Auf_alle_anwenden')?['Bezug']?['Value'],'X'),
              if(or(empty(triggerBody()?['GoLive']),empty(items('Auf_alle_anwenden')?['OffsetTage'])),null,
                 addDays(triggerBody()?['GoLive'],int(items('Auf_alle_anwenden')?['OffsetTage']),'yyyy-MM-dd')),
              null))
        ```
   3. *Element aktualisieren* – `DHO Häuser`, ID = Trigger-ID, `ChecklisteErstellt = Ja`

   **Nein-Zweig (Fristen neu berechnen, wenn Ü oder Go-Live geändert):**
   1. *Änderungen für ein Element abrufen* (nur Eigenschaften), Seit = `triggerOutputs()?['body/{TriggerWindowStartToken}']`
   2. Bedingung: `HasColumnChanged:Uebernahmetag` **oder** `HasColumnChanged:GoLive`
   3. *Elemente abrufen* – `DHO Aufgaben`, Filter `HausID eq <ID> and Status ne 'erledigt' and Status ne 'entfällt'`
   4. *Auf alle anwenden* → *Element aktualisieren*: Frist mit demselben Ausdruck wie oben

### Flow B: „DHO – Erledigt-Datum setzen“

- Trigger: *Wenn ein Element erstellt oder geändert wird*, Liste `DHO Aufgaben`
- Triggerbedingung (Einstellungen):
  `@and(equals(triggerBody()?['Status']?['Value'],'erledigt'),empty(triggerBody()?['ErledigtAm']))`
- Aktion: *Element aktualisieren* → `ErledigtAm = utcNow()`

### Flow C: „DHO – Wochenstatus“ (optional, passt zum PM-Takt)

- Montag 07:30: Mail an jeden Owner mit überfälligen und in 7 Tagen fälligen Punkten
- Freitag 12:00: Wochenstatus an die Projektleitung (Ansicht „Überfällig“ + „Nächste 14 Tage“ als HTML-Tabelle)

## 4. Dashboard live aus SharePoint

`index.html` liest die Listen `DHO Häuser`, `DHO Aufgaben` und `DHO Checkliste Vorlage`
direkt über Microsoft Graph, mit der Anmeldung des Benutzers (delegiert, nur lesen).
Es werden keine Daten im Repo oder auf einem Server gespeichert. Aktualisierung alle 5 Minuten
und per Knopf; jede Aufgabe verlinkt auf ihr SharePoint-Bearbeitungsformular.

Einmalige Einrichtung (Entra ID → App-Registrierungen → Neue Registrierung, z. B. „DHO-Dashboard“):

1. Kontotyp: nur dieses Organisationsverzeichnis
2. Plattform **Single-Page-Anwendung (SPA)**, Umleitungs-URIs:
   - `https://<github-user>.github.io/<repo>/`
   - `https://<github-user>.github.io/<repo>/index.html`
3. API-Berechtigungen → Microsoft Graph → Delegiert → `Sites.Read.All` → Administratorzustimmung erteilen
4. Dashboard aufrufen und Tenant-ID, Client-ID und Site-URL eintragen – oder als Lesezeichen:
   `https://<github-user>.github.io/<repo>/?tenant=<tenant-id>&client=<client-id>&site=https://<tenant>.sharepoint.com/sites/<site>`

Die Bibliothek `lib/msal-browser.min.js` (MSAL 2.39.0, MIT-Lizenz) liegt im Repo, damit die Seite ohne fremdes CDN anmeldet.
