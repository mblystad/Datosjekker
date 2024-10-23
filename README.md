# Excel Dato Kolonne Sammenligner

Et Python-basert GUI-program for å sammenligne dato-kolonner fra flere Excel-filer og rapportere eventuelle overlappende datoer. Programmet skriver resultatene til en oversiktelig tekstfil som kan brukes til videre analyse.

## Funksjonalitet
- **Velg flere Excel-filer**: Brukeren kan velge to eller flere Excel-filer som inneholder en kolonne med datoer som skal sammenlignes.
- **Dato-sammenligning**: Programmet sammenligner dato-kolonnene på tvers av de valgte filene og finner overlappende datoer.
- **Rapportgenerering**: Eventuelle overlappende datoer skrives til en rapportfil (`rapport_over_dato_kraesjer.txt`), der det spesifiseres hvilke datoer som overlapper mellom hvilke filer.
- **Visuell tilbakemelding**: Brukeren får en melding om at rapporten er lagret og hvor rapporten er lagret på disken.

## Krav
- Python 3.x
- Følgende Python-pakker:
  - `pandas`
  - `tkinter` (vanligvis inkludert i standard Python-distribusjonen)
  - `Pillow` (for å vise logoen)
  - `openpyxl` (for å lese Excel-filer)

Installer nødvendige pakker ved å kjøre:
```bash
pip install pandas pillow openpyxl
```

## Hvordan bruke

1. **Kjør programmet**: 
   - Kjør Python-filen som inneholder programmet. Dette vil åpne et GUI-vindu der du kan velge Excel-filer.
   
2. **Velg Excel-filer**:
   - Klikk på "Bla gjennom"-knappene for å velge Excel-filer med dato-kolonner. Programmet tillater deg å velge flere filer (standard satt til 3, men dette kan enkelt utvides i koden).

3. **Start sammenligningen**:
   - Klikk på "Sammenlign Datoer"-knappen. Programmet vil sammenligne dato-kolonnene fra filene.
   
4. **Resultater**:
   - Hvis det er overlappende datoer, vil disse skrives til en rapport i filen `rapport_over_dato_kraesjer.txt`. Brukeren får en melding om hvor rapporten er lagret.
   
## Rapportfil
Rapportfilen genereres som `rapport_over_dato_kraesjer.txt` i samme katalog som programmet. Filen vil inneholde:

- Hvis du sammenligner to filer, vil den vise hvilke datoer som kræsjer.
- Hvis du sammenligner flere filer, vil den vise hvilke datoer som kræsjer mellom hvilke filer.

Eksempel på innhold i rapportfilen:
```
Overlappende datoer mellom fil1.xlsx og fil2.xlsx:
Rad 1: 2024-01-15
Rad 2: 2024-02-28

Overlappende datoer mellom fil1.xlsx og fil3.xlsx:
Rad 1: 2024-03-12
Rad 2: 2024-04-19

Ingen overlappende datoer mellom fil2.xlsx og fil3.xlsx.
```

## Tilpasninger
- **Antall filer**: Du kan enkelt utvide antall filer som kan sammenlignes ved å justere `for i in range(3)`-loopen i GUI-koden. Øk verdien for å tillate flere filvalg.
- **Kolonnenavn**: Programmet antar at kolonnen med datoer heter "Dato". Hvis kolonnen har et annet navn, må du oppdatere dette i `compare_dates`-funksjonen i koden.

## Feilhåndtering
- Hvis det oppstår en feil under kjøringen (f.eks. feil filformat eller manglende kolonner), vil en feilmelding vises i et popup-vindu.

## Lisens
Dette prosjektet er lisensiert under MIT-lisensen. 

********MHB********
