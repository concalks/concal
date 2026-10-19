# Concal
An Outlook Add-in to add contact details to appointments.

## Development
- generate trusted local SSL certificates (either `mkcert` or `openssl`) and place them in the project root (`localhost.pem` and `localhost-key.pem`)
- run `npm run dev-server` to start the development server
- generate a new UUID for the add-in manifest and replace the placeholder in `manifest.xml` (necessary for Office to react to changes)
- sideload into OWA (Outlook Web App) via https://outlook.cloud.microsoft/mail/inclientstore (Careful: Microsoft could change the URL at any time)
