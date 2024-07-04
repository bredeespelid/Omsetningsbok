Dette skriptet behandler en Excel-fil ved å utføre ulike datamanipulasjoner og lagrer resultatet i en ny Excel-fil. Skriptet er brukervennlig og benytter GUI-dialoger for filvalg og lagringssted ved hjelp av Tkinter-biblioteket.

Forutsetninger
For å kjøre dette skriptet trenger du Python 3.x installert på systemet ditt sammen med følgende Python-pakker: pandas, openpyxl, xlsxwriter, og tkinter.

Funksjonalitet
Skriptet utfører følgende operasjoner:

Filvalg: Skriptet bruker en filvelger-dialog for å be brukeren velge en Excel-fil (.xlsx).
Les Excel-fil: Den valgte Excel-filen leses inn i en DataFrame, der de første 7 radene hoppes over.
Legg til 'Avd' Kolonne: En ny kolonne med navnet 'Avd' legges til venstre for 'Navn'-kolonnen. Skriptet henter ut numeriske verdier fra 'Navn'-kolonnen og fyller 'Avd'-kolonnen med disse verdiene, og viderefører den siste gyldige verdien fremover. Eventuelle gjenværende NaN-verdier i 'Avd'-kolonnen fylles med 0, og kolonnen konverteres til heltall.
Filtrer Rader Basert på 'Konto' Kolonne: Skriptet definerer en funksjon for å filtrere ut rader basert på spesifikke kriterier i 'Konto'-kolonnen og anvender denne funksjonen på DataFrame.
Konverter Data til Heltall: Skriptet konverterer celler nedenfor og til høyre for D1 til heltall der det er mulig.
Konverter Kolonnenavn til Datoformat: Kolonnenavn fra D1 og videre konverteres fra 'dd.mm.yyyy'-format til 'yyyy/mm/dd'-format.
Filtrer Rader Basert på 'Avd' Kolonne: Skriptet filtrerer ut rader med tallene 10 og 90 i 'Avd'-kolonnen.
Lagre Bearbeidede Data til Ny Excel-fil: Brukeren blir bedt om å velge et lagringssted for den nye Excel-filen. DataFrame lagres deretter i den nye Excel-filen, med DATEVALUE-formler lagt til i kolonneoverskriftene for kolonner med datoer.
Suksessmelding: En suksessmelding vises når filen er lagret.

Bruk Kjør Skriptet: Kjør skriptet i et Python-miljø.
Velg Inndatafil: En filvelger-dialog vil dukke opp. Velg Excel-filen (.xlsx) du ønsker å behandle.
Lagre Utdatapfilen: Etter behandling vil en annen filvelger-dialog dukke opp. Velg hvor du vil lagre den nye Excel-filen.
Fullføring: En meldingsboks bekrefter at filen er lagret.
