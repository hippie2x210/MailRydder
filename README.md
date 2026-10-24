# MailRydder

MailRydder rydder e-post i Gmail eller Outlook/Exchange. Det viser først en forhåndsvisning, og meldinger som ryddes flyttes til Papirkurv/Slettede elementer.

## Start programmet

1. Last ned ZIP-filen fra GitHub og pakk den ut.
2. Åpne mappen `MailRydder-master`, og åpne den indre mappen `MailRydder-master` hvis den finnes.
3. Dobbeltklikk **Start-MailRydder.bat**.
4. Velg Gmail eller Exchange/Outlook i vinduet.
5. Velg **Forhåndsvisning** først for å se hva programmet finner.

Du trenger ikke en `.exe`-fil. Ikke start `Test-Exchange.bat` eller `Test-MailRydder.bat` for vanlig bruk; de er testfiler.

## Exchange/Outlook første gang

Velg Exchange/Outlook og start en forhåndsvisning. Når Microsoft-innlogging åpnes, logg inn på kontoen som skal ryddes og godkjenn MailRydder. Programmet ber aldri om Microsoft-passordet inne i appen.

## Trygghet

- Forhåndsvisning flytter eller sletter ingenting.
- Rydding flytter e-post til Papirkurv/Slettede elementer, ikke permanent sletting.
- Beskyttede meldinger og flaggede/stjernemerkede meldinger beholdes.
- `Test-Exchange.bat` og `Test-MailRydder.bat` kjører bare testmodus.
