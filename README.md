# Begrepsliste → Anki

Statisk nettside som gjør en utfylt begrepsliste (.docx, eller innlimt fra Word/Google Docs) om til en Anki-pakke (.apkg). Alt skjer i studentens nettleser; ingen filer lastes opp.

## Publisere på GitHub Pages (ca. 5 min)

1. Lag et nytt, offentlig repo på github.com, f.eks. `begrepsliste-anki`.
2. «Add file → Upload files»: last opp **innholdet** i denne mappen (`index.html`, `konverter.js`, mappen `lib/`).
3. Settings → Pages → Source: «Deploy from a branch», branch `main`, mappe `/ (root)` → Save.
4. Etter et minutt ligger siden på `https://<brukernavn>.github.io/begrepsliste-anki/`. Del lenken på Blackboard/Canvas.

Test lokalt: `python3 -m http.server` i denne mappen, åpne http://localhost:8000 (må kjøres via en server, ikke dobbeltklikk – sql.js laster en .wasm-fil).

## Tilpasning

- Standard kortstokknavn: `value="EXPH0400 H26"` i `index.html`. Endre til f.eks. `EXPH0400 V27` neste semester.
- Kortene får stabile ID-er ut fra begrepsteksten + kortstokknavnet. Studenter kan derfor importere på nytt når de har fylt ut mer: eksisterende kort oppdateres, nye legges til, og repetisjonshistorikken beholdes. Endrer man ordlyden i selve begrepet (kolonne 1), blir det et nytt kort.
- Malens hjelpetekster («(beskriv forskjellen mellom disse to)», «1.–4.», «nødvendig/tilstrekkelig/eksempler:») regnes som tomme og blir ikke kort. Listen står i `PLACEHOLDER_LINES` i `konverter.js`.

## Kortformat

Notetype «Begrepsliste (begrep → definisjon)» med feltene Begrep, Definisjon, Kilde, Tema og Baklengs.
Kort 1: begrep → definisjon + kilde. Kort 2 (kun hvis «baklengs-kort» er krysset av): definisjon med begrepet skjult som […] → begrep.
Én underkortstokk per ukeoverskrift, f.eks. `EXPH0400 H26::01 Argumentasjonsteori (uke 35)`, og tagger `EXPH0400` + `uke35`.

Testet med Anki 26.8.1 (import, gjenimport uten duplikater, oppdatering av endrede kort).

Biblioteker: JSZip 3.10.1 (MIT/GPLv3), sql.js 1.10.3 (MIT) – se `lib/`.
