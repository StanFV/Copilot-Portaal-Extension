/**
 * Het paneel in Outlook: een bibliotheek met Copilot-prompts. De lijst staat in prompts.json en
 * komt uit het Copilot Portaal (scripts/outlook-export.ps1 in de portaalrepo). Het paneel leest de
 * mail niet; alleen de favorieten worden bewaard, in de Outlook-instellingen van de gebruiker.
 */
import './stijl.css';

interface Prompt {
  id: string;
  rubriek: string;
  titel: string;
  prompt: string;
  /** Waar en hoe je de prompt gebruikt. */
  uitleg?: string;
  /** Wat de prompt oplevert (opdrachten uit het traject). */
  omschrijving?: string;
  /** Alleen bij opdrachten uit het traject. */
  week?: number;
}

interface Lijst {
  bijgewerkt: string;
  set: string;
  rubrieken: string[];
  prompts: Prompt[];
}

const ALLE = 'Alle';
const FAVORIETEN = 'Favorieten';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const inhoud = $<HTMLElement>('inhoud');
const rubriekBalk = $<HTMLElement>('rubrieken');
const zoekVeld = $<HTMLInputElement>('zoek');
const toast = $<HTMLElement>('toast');

let lijst: Lijst | null = null;
let rubriek = ALLE;
let zoek = '';
let favorieten = new Set<string>();
let uitlegGezien = false;

/* ------------------------------------------------------------------ opslag */

/** In Outlook: de roaming settings (gaan mee naar andere apparaten). Daarbuiten: localStorage. */
const opslag = {
  inOutlook: () => typeof Office !== 'undefined' && Boolean(Office.context?.roamingSettings),
  lees<T>(sleutel: string): T | undefined {
    try {
      if (this.inOutlook()) return (Office.context.roamingSettings.get(sleutel) as T) ?? undefined;
      const s = localStorage.getItem(`bru-outlook-${sleutel}`);
      return s ? (JSON.parse(s) as T) : undefined;
    } catch {
      return undefined;
    }
  },
  schrijf(sleutel: string, waarde: unknown) {
    try {
      if (this.inOutlook()) {
        Office.context.roamingSettings.set(sleutel, waarde);
        Office.context.roamingSettings.saveAsync();
      } else {
        localStorage.setItem(`bru-outlook-${sleutel}`, JSON.stringify(waarde));
      }
    } catch {
      /* niet bewaard; het paneel werkt gewoon verder */
    }
  },
};

/* ------------------------------------------------------------------ hulp */

function escape(t: string) {
  return t.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
}

/** Prompttekst met de invulplekken [zo] gemarkeerd. */
function promptHtml(t: string) {
  return escape(t).replace(/\[[^\]]+\]/g, (m) => `<mark class="invul">${m}</mark>`);
}

/** Naar het klembord; eerst de moderne manier, anders de oude (sommige Outlook-versies). */
async function kopieer(tekst: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(tekst);
    return true;
  } catch {
    const ta = document.createElement('textarea');
    ta.value = tekst;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try {
      ok = document.execCommand('copy');
    } catch {
      ok = false;
    }
    ta.remove();
    return ok;
  }
}

let toastTimer = 0;
function meld(tekst: string) {
  toast.textContent = tekst;
  toast.hidden = false;
  toast.classList.remove('weg');
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => {
    toast.classList.add('weg');
    toastTimer = window.setTimeout(() => (toast.hidden = true), 250);
  }, 3200);
}

/* ------------------------------------------------------------------ weergave */

function zichtbare(): Prompt[] {
  if (!lijst) return [];
  const t = zoek.trim().toLowerCase();
  return lijst.prompts.filter(
    (p) =>
      (rubriek === ALLE || (rubriek === FAVORIETEN ? favorieten.has(p.id) : p.rubriek === rubriek)) &&
      (!t || `${p.titel} ${p.prompt} ${p.rubriek} ${p.omschrijving ?? ''}`.toLowerCase().includes(t)),
  );
}

function tekenRubrieken() {
  if (!lijst) return;
  const knoppen = [ALLE, ...(favorieten.size ? [FAVORIETEN] : []), ...lijst.rubrieken];
  if (!knoppen.includes(rubriek)) rubriek = ALLE;
  rubriekBalk.innerHTML = knoppen
    .map((r) => `<button type="button" class="chip${r === FAVORIETEN ? ' chip-fav' : ''}" aria-pressed="${r === rubriek}" data-rubriek="${escape(r)}">${r === FAVORIETEN ? '★ ' : ''}${escape(r)}</button>`)
    .join('');
  rubriekBalk.querySelector<HTMLElement>('[aria-pressed="true"]')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  randRubrieken();
}

/** Vervaagde rechterrand zolang er rechts nog rubrieken staan. */
function randRubrieken() {
  const b = rubriekBalk;
  b.classList.toggle('meer', b.scrollLeft + b.clientWidth < b.scrollWidth - 20);
}

function kaart(p: Prompt) {
  const fav = favorieten.has(p.id);
  return `
    <article class="kaart" data-id="${escape(p.id)}">
      <div class="kaart-kop">
        <h3>${escape(p.titel)}</h3>
        <button type="button" class="ster" data-fav="${escape(p.id)}" aria-pressed="${fav}" aria-label="${fav ? 'Verwijder uit favorieten' : 'Bewaar als favoriet'}" title="${fav ? 'Verwijder uit favorieten' : 'Bewaar als favoriet'}">${fav ? '★' : '☆'}</button>
      </div>
      ${p.week !== undefined ? `<span class="label">${p.week === 0 ? 'Kick-off' : `Week ${p.week}`} van het traject</span>` : ''}
      ${p.omschrijving ? `<p class="omschrijving">${escape(p.omschrijving)}</p>` : ''}
      <p class="prompt">${promptHtml(p.prompt)}</p>
      ${p.uitleg ? `<details class="hoe"><summary>Waar plak ik dit?</summary><p>${escape(p.uitleg)}</p></details>` : ''}
      <button type="button" class="knop" data-kopieer="${escape(p.id)}">
        <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><rect x="8" y="8" width="12" height="12" rx="2" /><path d="M16 8V5a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1v10a1 1 0 0 0 1 1h3" /></svg>
        Kopieer prompt
      </button>
    </article>`;
}

function uitlegBlok() {
  if (uitlegGezien) return '';
  return `
    <section class="uitleg" aria-label="Zo werkt het">
      <h2>Zo werkt het</h2>
      <ol>
        <li>Kies een prompt en klik op <strong>Kopieer prompt</strong>.</li>
        <li>Open Copilot in Outlook: rechtsboven, of in je mail via het Copilot-logo.</li>
        <li>Plak met <strong>Ctrl+V</strong> en vul de <mark class="invul">[blauwe stukjes]</mark> zelf in.</li>
      </ol>
      <button type="button" class="knop stil" data-uitleg-weg>Begrepen</button>
    </section>`;
}

function teken() {
  tekenRubrieken();
  if (!lijst) return;
  const items = zichtbare();
  let html = uitlegBlok();
  if (!items.length) {
    html += `<p class="leeg">${rubriek === FAVORIETEN ? 'Je hebt nog geen favorieten. Klik op de ster bij een prompt om hem hier te bewaren.' : `Geen prompts gevonden${zoek.trim() ? ` voor ‘${escape(zoek.trim())}’` : ''}.`}</p>`;
  } else if (rubriek === ALLE && !zoek.trim()) {
    for (const r of lijst.rubrieken) {
      const groep = items.filter((p) => p.rubriek === r);
      if (groep.length) html += `<section class="groep"><h2>${escape(r)}</h2>${groep.map(kaart).join('')}</section>`;
    }
  } else {
    html += items.map(kaart).join('');
  }
  inhoud.innerHTML = html;
}

/* ------------------------------------------------------------------ gebeurtenissen */

rubriekBalk.addEventListener('click', (e) => {
  const knop = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-rubriek]');
  if (!knop) return;
  rubriek = knop.dataset.rubriek!;
  teken();
  inhoud.scrollTop = 0;
  window.scrollTo({ top: 0 });
});

rubriekBalk.addEventListener('scroll', randRubrieken, { passive: true });
window.addEventListener('resize', randRubrieken);
// Met het muiswiel horizontaal door de rubrieken (klassieke Outlook heeft vaak geen touchpad).
rubriekBalk.addEventListener(
  'wheel',
  (e) => {
    if (Math.abs(e.deltaY) <= Math.abs(e.deltaX) || rubriekBalk.scrollWidth <= rubriekBalk.clientWidth) return;
    rubriekBalk.scrollLeft += e.deltaY;
    e.preventDefault();
  },
  { passive: false },
);

zoekVeld.addEventListener('input', () => {
  zoek = zoekVeld.value;
  teken();
});

inhoud.addEventListener('click', async (e) => {
  const doel = e.target as HTMLElement;
  const kopieerKnop = doel.closest<HTMLButtonElement>('[data-kopieer]');
  if (kopieerKnop) {
    const p = lijst?.prompts.find((x) => x.id === kopieerKnop.dataset.kopieer);
    if (!p) return;
    if (await kopieer(p.prompt)) {
      kopieerKnop.classList.add('gekopieerd');
      kopieerKnop.lastChild!.textContent = ' Gekopieerd';
      setTimeout(() => {
        kopieerKnop.classList.remove('gekopieerd');
        kopieerKnop.lastChild!.textContent = ' Kopieer prompt';
      }, 2000);
      meld(/\[[^\]]+\]/.test(p.prompt) ? 'Gekopieerd. Plak in Copilot met Ctrl+V en vul de blauwe stukjes in.' : 'Gekopieerd. Plak in Copilot met Ctrl+V.');
    } else {
      // Kopiëren mag hier niet: selecteer de tekst zodat de gebruiker zelf Ctrl+C kan doen.
      const tekst = kopieerKnop.closest('.kaart')?.querySelector('.prompt');
      if (tekst) window.getSelection()?.selectAllChildren(tekst);
      meld('Kopiëren lukte niet. De tekst is geselecteerd: druk op Ctrl+C.');
    }
    return;
  }
  const ster = doel.closest<HTMLButtonElement>('[data-fav]');
  if (ster) {
    const id = ster.dataset.fav!;
    if (favorieten.has(id)) favorieten.delete(id);
    else favorieten.add(id);
    opslag.schrijf('favorieten', [...favorieten]);
    meld(favorieten.has(id) ? 'Bewaard bij je favorieten.' : 'Verwijderd uit je favorieten.');
    teken();
    return;
  }
  if (doel.closest('[data-uitleg-weg]')) {
    uitlegGezien = true;
    opslag.schrijf('uitlegGezien', true);
    teken();
  }
  if (doel.closest('[data-opnieuw]')) laad();
});

/* ------------------------------------------------------------------ start */

async function laad() {
  inhoud.innerHTML = '<p class="laden">Prompts laden…</p>';
  try {
    const r = await fetch('./prompts.json', { cache: 'no-cache' });
    if (!r.ok) throw new Error(String(r.status));
    lijst = (await r.json()) as Lijst;
    // Favorieten van prompts die niet meer bestaan, vallen weg.
    favorieten = new Set([...favorieten].filter((id) => lijst!.prompts.some((p) => p.id === id)));
    teken();
  } catch {
    inhoud.innerHTML = '<div class="leeg"><p>De prompts konden niet worden geladen. Controleer je internetverbinding.</p><button type="button" class="knop stil" data-opnieuw>Opnieuw proberen</button></div>';
  }
}

let gestart = false;
function start() {
  if (gestart) return;
  gestart = true;
  favorieten = new Set(opslag.lees<string[]>('favorieten') ?? []);
  uitlegGezien = opslag.lees<boolean>('uitlegGezien') ?? false;
  laad();
}

// Zonder office.js (voorbeeld.html op de installatiepagina) meteen starten. Anders wachten op
// Outlook, met een vangnet voor als Office niet antwoordt.
if (typeof Office === 'undefined') start();
else {
  Office.onReady(() => start());
  setTimeout(start, 5000);
}
