# Copilot-Portaal-Extension: Bruinsma Copilot-prompts voor Outlook

Outlook-invoegtoepassing van Bruinsma Kantoor en ICT. Een knop **Copilot-prompts** in het lint van Outlook opent een paneel met kant-en-klare prompts voor Microsoft 365 Copilot: samenvatten, beantwoorden, opstellen, agenda en vergaderingen, opvolgen en plannen, toon en nalezen, plus de Outlook-opdrachten uit het Copilot-traject.

- **Installatiepagina:** https://stanfv.github.io/Copilot-Portaal-Extension/
- **Manifest:** https://stanfv.github.io/Copilot-Portaal-Extension/manifest.xml

## Hoe het werkt

- `taskpane.html` + `src/taskpane.ts`: het paneel in Outlook. Leest `prompts.json`, met zoeken, rubrieken, favorieten en een knop om de prompt te kopiëren. Het paneel leest geen mail (rechten `Restricted`). Favorieten staan in de roaming settings van Outlook (buiten Outlook in `localStorage`).
- `index.html`: de installatiepagina, met rechts een live voorbeeld (`voorbeeld.html`, door `vite.config.ts` gemaakt uit `taskpane.html` zonder office.js: office.js stuurt een iframe buiten Office naar `about:blank`).
- `public/manifest.xml`: XML-manifest (Mailbox 1.1, knoppen bij mail lezen en schrijven en bij afspraken, vast te zetten paneel). Verhoog `<Version>` bij elke wijziging aan het manifest.
- `public/prompts.json`: de prompts. **Niet met de hand bewerken.** Het bestand komt uit het Copilot Portaal (privé-repo copilot-portaal): beheer de prompts daar op de pagina **Outlook** en draai `./scripts/outlook-export.ps1 -Publiceer`. Dat schrijft dit bestand, commit en pusht.
- `docs/`: de gebouwde site (`npm run build`). GitHub Pages serveert deze map vanaf `main`. Bouw en commit `docs/` mee na elke wijziging; `outlook-export.ps1` doet dat zelf.

## Ontwikkelen

```bash
npm install
npm run dev      # http://localhost:5173/taskpane.html
npm run build    # naar docs/ (wordt gepubliceerd)
npx office-addin-manifest validate public/manifest.xml
```

## Installeren

Zie de installatiepagina. Kort: in het Microsoft 365-beheercentrum > Instellingen > Geïntegreerde apps > Aangepaste apps uploaden > Office-invoegtoepassing, met de link naar het manifest. Of zelf via https://aka.ms/olksideload > Mijn invoegtoepassingen > Aangepaste invoegtoepassing toevoegen > Uit bestand toevoegen.

Vragen: help@bruinsma-ict.nl
