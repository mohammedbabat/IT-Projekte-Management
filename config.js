// Standard-Verbindung des DHO-Dashboards.
// Tenant-ID und Client-ID sind keine Geheimnisse (öffentliche Anmelde-App ohne Secret):
// Zugriff haben nur Mitglieder der Gruppe, die der App in Entra ID zugewiesen ist,
// und Daten sieht nur, wer Rechte auf der SharePoint-Site hat.
// Ein Link mit ?tenant=&client=&site=&mail= oder die Verbindungseinstellungen überschreiben diese Werte.
window.DHO_CONFIG = {
  tenant: '69a7d3e4-6d53-43dd-a016-dc755b1f684b',
  client: '4c131eb5-395e-4cbc-8c6d-66667e2f03b1',
  site: 'https://6zxkpr.sharepoint.com/sites/IT-Projekte-Management',
  mail: 'IT-Projekt@ein-facher.de'
};
