# Sprint Timer

Mobilní měření sprintu dvojic se zápisem do Google Sheets. Bez vlastního backendu –
statické `index.html` + Google Apps Script + Google Sheet.

## Architektura

```
Mobil / prohlížeč
   → index.html (měření + outbox fronta v localStorage)
   → HTTP POST (text/plain, no-cors)
   → Google Apps Script Web App (zámek, kontrola tokenu, idempotence)
   → Google Sheet: Startovka (vstup) + Results (výstup)
```

## Co řešení dělá oproti naivní verzi

- **`no-cors` + `text/plain`** – Apps Script `/exec` neposílá CORS hlavičky a redirectuje;
  v běžném `cors` módu by `fetch` spadl i při úspěšném zápisu. Proto posíláme „naslepo"
  a správnost zápisu řešíme na serveru.
- **Outbox v localStorage** – při výpadku signálu se jízda nezahodí; appka ji drží ve
  frontě a posílá automaticky každých 15 s a při návratu signálu.
- **Idempotence (`id` jízdy)** – každá jízda má unikátní ID; server podruhé stejné ID
  nezapíše, takže retry netvoří duplicity.
- **`LockService`** – serializuje souběžné zápisy z více mobilů.
- **Sloupec `Čas (ms)`** – surové milisekundy pro řazení a leaderboard přímo v Sheetu.
- **Sdílený token** – základní pojistka proti náhodnému spamu endpointu.

## Nastavení

### 1. Google Sheet

Vytvoř Sheet a list **`Startovka`** s hlavičkou a daty:

| Startovní číslo | Jméno | Kategorie |
|---|---|---|
| 101 | Jan Novák | 6–8 let |
| 102 | Petr Dvořák | 6–8 let |

List **`Results`** se vytvoří automaticky při prvním zápisu (i s hlavičkou).

### 2. Apps Script

1. V Sheetu: **Rozšíření → Apps Script**, vlož obsah `apps-script.gs`, ulož.
2. Nastav `SHARED_SECRET` na vlastní tajný řetězec.
3. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Povol oprávnění, zkopíruj `/exec` URL.

> Po každé další změně kódu: **Manage deployments → edit → Version: New version**,
> jinak běží stará verze.

### 3. Frontend

V `index.html` vyplň:

```javascript
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/XXXXX/exec";
const SHARED_SECRET     = "stejny-retezec-jako-v-apps-scriptu";
```

Hostuj na GitHub Pages / Netlify / Vercel, nebo otevři lokálně.

### 4. Kontrola dat

Kdokoli se sdíleným přístupem do Google Sheetu (klidně jen pro čtení) vidí list
`Results` živě a může ručně opravovat (např. neznámá čísla, sloupec lze rozšířit
o `Jízda`).

## Test

1. Otevři `doGet` URL v prohlížeči → měl by vrátit `{"status":"OK",...}`.
2. V appce: START → STOP 1 → STOP 2 → zadej dvě čísla → ODESLAT.
3. Zkontroluj dva nové řádky v `Results`.
4. Test offline: vypni síť, odešli – jízda zůstane „Čeká na odeslání"; zapni síť →
   automaticky se doručí.

## Známá omezení

- Token je ve statickém HTML viditelný – chrání jen proti náhodnému/botímu spamu,
  ne proti cílenému útoku. Pro malý závod dostačuje.
- Měření používá `Date.now()` klienta; přesnost odpovídá hodinám zařízení (na sprint
  v řádu sekund/minut bohatě stačí, časy se počítají z rozdílu, ne z absolutního času).
