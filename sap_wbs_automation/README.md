# Automatizare SAP: WBS nivel 3/4 + norme din devizul general

Scriptul citește devizul general (`.xlsm`), deduce structura WBS (nivel 3 și 4), activitățile din
rețelele standard și normele (serviciile) cu cantitățile lor. Apoi le introduce în SAP prin
SAP GUI Scripting (`win32com.client`), în tranzacția de proiect (implicit `CJ02`).

```
Proiect ─┬─ +M1  MEDIE TENSIUNE                       (nivel 3)
         │    ├─ +441R1  CLADIRE PT1   [rețea A441]    (nivel 4)
         │    │     └─ YA1Furniz/ const/ achiz post prefabr (L)   (activitate)
         │    │           ├─ L31104 = 1                           (servicii / norme)
         │    │           └─ L34103 = 1,2
         │    └─ +453R4  RETEA LES MT COMUN  [rețea A453]
         ├─ +B1  JOASA TENSIUNE
         └─ +T1  TELECONTROL
```

## Instalare (Windows, pe stația cu SAP GUI)

1. Python 3.9 sau mai nou.
2. `pip install -r requirements.txt`
3. Activați SAP GUI Scripting:
   - în SAP GUI: *Options → Accessibility & Scripting → Scripting → Enable scripting*;
     debifați notificările „When a script attaches / opens a connection”;
   - pe server, parametrul `sapgui/user_scripting = TRUE` (RZ11; îl setează echipa Basis).
4. Deschideți SAP Logon și autentificați-vă. Scriptul se atașează la o sesiune existentă,
   de preferință una aflată în meniul principal.

## Utilizare

```bat
REM 1. Doar citire + raport pre-flight (nu atinge SAP). Rulați întotdeauna întâi asta.
python run.py "C:\Devize\NEC_v20.xlsm" --dry-run

REM 2. Primul test în SAP: un singur element de nivel 4, cu confirmare înainte
python run.py "C:\Devize\NEC_v20.xlsm" --max-n4 1 --pas-cu-pas

REM 3. Rulare completă
python run.py "C:\Devize\NEC_v20.xlsm"
```

| Opțiune | Rol |
|---|---|
| `--dry-run` | citește devizul, afișează arborele și scrie `rapoarte/plan_*.xlsx` |
| `--proiect C-0001` | definiția de proiect (altfel se citește din Excel) |
| `--cod-buget REM_001604` | codul de buget (altfel se citește din Excel) |
| `--max-n4 N` | procesează cel mult N elemente de nivel 4 |
| `--pas-cu-pas` | cere confirmare (Da / Nu / Stop) înainte de fiecare element de nivel 4 |
| `--recalculeaza` | recalculează formulele în Excel înainte de citire (dacă fișierul nu a fost salvat de Excel) |
| `--continua-cu-avertizari` | rulează în SAP chiar dacă planul are erori (de ex. lipsește codul de buget) |
| `--config alt.yaml` | alt fișier de configurare |

Codurile de ieșire: `0` = totul OK; `1` = erori (de plan sau la unele rânduri); `2` = fișier sau
configurație invalidă; `3` = conexiunea SAP a eșuat.

## Cum este dedusă structura devizului

Nicio coloană nu este indexată fix. Totul se configurează în `config.yaml`.

1. **Detectarea antetului.** În fiecare foaie vizibilă se scanează primele 40 de rânduri. Fiecare
   celulă este comparată cu alias-urile din `excel.columns`, fără să conteze majusculele, diacriticele
   sau punctuația. Potrivirea exactă valorează 100 de puncte, „începe cu” 60, „conține cuvântul” 30.
   Antetul este rândul cu scorul maxim care conține coloanele obligatorii (cod, cantitate, rețea).
   Foile fără ele sunt ignorate. De exemplu, în foile „Norme” antetul e pe rândul 17
   (`Cod Norma`, `Cant.`, coloana AI). În foile „MAT” e pe rândul 13 (`Cod normă`, `Cantitate totala`,
   coloana AH). Ambele sunt găsite automat.
2. **Validarea după conținut.** Coloana de cod este verificată cu regex `^[A-Z]\d{5}$`. Dacă antetul
   indică greșit coloana, se alege cea cu cele mai multe coduri valide.
3. **Rândurile.** Se păstrează doar rândurile cu cod valid și cantitate nenulă. Subtotalurile și
   titlurile cad singure.
4. **Partea (sufixul activității).** Vine din coloana `Partea` (L / M / S). La foile de materiale,
   unde coloana lipsește, se ia prima literă a codului (E). Părțile I, W și T sunt ignorate.
5. **Rețeaua standard și activitatea.** Vin din `Activitate SAP` (A453) și
   `Denumire activitate SAP` (XL4 …). Când există variante (`A451/A452`, `YL1 … / YM1 …`), regulile
   se aplică în ordine:
   1. coloana opțională `Activitate SAP ales`, dacă o adăugați (forțează alegerea, ex. `A452` sau `A452 XA3`);
   2. perechile poziționale (rețeaua i ↔ activitatea i);
   3. contextul din foaia Date: *Nivel Tensiune* (MT/JT) și *Soluție proiectată* (LEA/LES);
   4. rețelele deja folosite neambiguu în același deviz (ex. un material XD2 implică A451);
   5. cuvintele cheie din descrierea normei (demontare / înlocuire / montare);
   6. prima variantă, cu avertizare. Puneți `ambiguity_fallback: skip` ca să sară rândul în loc.

   Orice alegere euristică apare ca **avertizare** în raportul pre-flight.
6. **Nivelul 3.** Se citește din tabelul „Activitate SAP / Nivel 3” din foaia Date. Dacă tabelul
   lipsește, se folosește `rules.networks`.
7. **Nivelul 4 (convenția aleasă).** Fiecare rețea are propriile elemente de nivel 4:
   - cantitatea din coloana **PTk** merge pe `+{rețea}R{k}`, cu denumirea „{bază} PT{k}”
     (ex. `+441R2 CLADIRE PT2`);
   - restul (cantitate totală − suma PT) înseamnă **lucrări comune**:
     - dacă rețeaua nu are nimic pe PT, elementul este `R1`, ca în înregistrarea SAP
       (ex. `+452R1 LEA MT COND IZOLAT`);
     - altfel primește primul indice liber după ultimul PT (cu 3 PT-uri, `+453R4 RETEA LES MT COMUN`).

   Rezultatul pe devizul NEC_v20 coincide cu calculul „Nr. nivele 4 necesare” din foaia Date
   (A412 × 4, A414 × 3, A441 × 4, A453 × 4, A463 × 4, A61F × 4).
8. **Activitatea în SAP.** Textul este denumirea fără unitatea de măsură finală (`NUM`), plus sufixul.
   De exemplu „XN4 Mans/termin Linii MT cablu subt NUM” devine „XN4 Mans/termin Linii MT cablu subt (L)”,
   tăiat la 40 de caractere. Normele cu același cod în aceeași activitate se adună.
9. **Controlul de sumă.** Suma cantităților citite trebuie să fie egală cu suma planificată.
   Diferențele apar în raport.

## Celula pentru codul de buget (și definiția de proiect)

Codul de buget ajunge în câmpul `PRPS-ZZCOD_BUGET` al fiecărui element de nivel 4. Când celula va
exista în Excel, oricare dintre variantele de mai jos funcționează, fără nicio modificare de cod:

- **Nume definit** (recomandat): selectați celula și dați-i numele `COD_BUGET`
  (*Formulas → Name Manager*). Funcționează oriunde ar fi celula.
- **Etichetă:** scrieți `Cod buget` într-o celulă din foaia `Date` și valoarea în celula din dreapta.
- **Adresă fixă:** în `config.yaml`, `wbs.budget_code.cell: "Date!C7"`.

Definiția de proiect folosește același mecanism: numele `COD_PROIECT` sau eticheta `Cod proiect SAP`.
Alternativ, o dați cu `--proiect`. Fără cod de buget sau fără proiect, scriptul nu atinge SAP-ul.

## Ce face în SAP (tradus din `Script1.vbs`)

Pentru fiecare element de nivel 3, apoi pentru fiecare element de nivel 4:

1. `/nCJ02`, definiția de proiect, apoi lista elementelor WBS (`ProjectBuilder.open_project`).
2. Dacă elementul nu există, `create_wbs_element(parent, cod, nivel, denumire)`: rând nou cu nivel,
   cod și denumire. Se verifică întâi că rândul nou ajunge sub părintele corect.
3. Atribuirile (`STORT`, `TPLNR` sau `EQUNR`) și câmpurile client (`ZZCOD_BUGET`, `ZZPARTE`), apoi salvare.
4. Activitățile:
   1. *copiere rețea standard* (`STDNR` = rețeaua, `STDAL` = `QS`);
   2. se păstrează **după text** (codul de 3 caractere, nu după poziție) activitățile necesare;
      restul se șterg.
   3. Dacă aceeași activitate e necesară de mai multe ori (L, M, E), rețeaua se copiază de mai multe ori.
5. Pentru fiecare activitate, `add_norm_to_activity(...)`:
   1. textul activității, `EKGRP` și `AFNAM` = `AMC`;
   2. *Servicii*: cod serviciu, cantitate (format SAP `348,6`), Enter, bifa „nelimitat”;
   3. serviciile valorice `E00xxx` primesc preț brut 1.
6. Salvare, apoi următorul element.

Re-rularea este sigură: elementele WBS, activitățile (după text) și serviciile (după cod) care
există deja nu se mai creează. Apar ca `EXISTENT` în raport.

## Erori, log și rapoarte

- Logul se scrie în `logs/sap_wbs_<data>.log` (DEBUG) și în consolă (INFO), cu timestamp.
- **Eroare la o normă** (cod inexistent, format greșit): rândul se anulează în SAP, mesajul din bara
  de stare se loghează, iar bucla continuă. Mesajul arată așa:
  `Eroare la WBS +453R4, rândul Norme MT!398 (norma L22101, activitatea 'XL4 … (L)'): <mesaj SAP>`
- **Eroare la o activitate:** se revine la lista de activități și se continuă cu următoarea.
- **Eroare la un element WBS:** modificările nesalvate se abandonează (`/n`) și se trece la
  următorul element. După 3 eșecuri consecutive execuția se oprește, de exemplu dacă sesiunea s-a închis.
- **Pop-up-uri:**
  - cererile de valori (AMC, MBT1 la crearea rețelei) se completează din `sap.network_popup_values`;
  - confirmările Da/Nu primesc răspuns automat;
  - jurnalul de erori la salvare se raportează ca eroare;
  - informările se închid cu Enter.

  Textul fiecărui pop-up ajunge în log.
- **Rapoarte:**
  - `rapoarte/plan_<data>.xlsx` (pre-flight): foile *Plan*, *Probleme*, *Sumar*, *Coloane detectate*;
  - `rapoarte/rezultate_<data>.xlsx`: statusul fiecărei norme (OK / EXISTENT / EROARE / BLOCAT /
    NEPROCESAT) și mesajul SAP.

## De verificat la prima rulare în SAP

ID-urile GUI sunt toate în `sapwbs/sap_ids.py`. Cele din înregistrare sunt sigure. Următoarele nu
apar în înregistrare, deci verificați-le cu `--max-n4 1 --pas-cu-pas`. Dacă diferă, o înregistrare
scurtă a pasului respectiv e suficientă pentru corectură.

| Ce | Unde | Observație |
|---|---|---|
| Câmpul „Definiție proiect” pe ecranul inițial | `INITIAL_PROJECT_FIELD = wnd[0]/usr/ctxt*PROJ-PSPID` | neînregistrat |
| Butonul spre lista WBS | `BTN_WBS_OVERVIEW = wnd[0]/tbar[1]/btn[5]` | dedus din înregistrare (folosit după salvare) |
| Închiderea detaliului de activitate | `ACT_DETAIL_CLOSE_BTN = wnd[1]/tbar[0]/btn[5]` | apăsat după F3 din servicii |
| Formatul codului WBS în listă | `cj02.ident_matches` | acceptă `+453R1` și `C-0001-453R1` |

Limitări cunoscute:

- Ierarhia din lista WBS e dată de ordine + nivel, iar elementele noi se adaugă la final. Un nivel 4
  nou poate fi deci creat doar sub **ultimul** nivel 3 din listă, sau sub unul creat în aceeași
  rulare. Celelalte cazuri sunt raportate ca `BLOCAT` și nu se creează nimic greșit.
- Activitățile interne aduse de rețeaua standard (tab-ul intern) nu sunt șterse. Înregistrarea
  procedează la fel.

## Structura proiectului

```
sap_wbs_automation/
├── run.py                 # punctul de intrare (python run.py -h)
├── config.yaml            # alias-uri coloane, reguli, convenții WBS, parametri SAP
├── sapwbs/
│   ├── config.py          # configurația (dataclasses + încărcare YAML validată)
│   ├── excel_reader.py    # detectarea antetelor, extragerea normelor, contextul din foaia Date
│   ├── rules.py           # motorul de reguli: rețea, activitate, sufix, denumiri WBS
│   ├── planner.py         # arborele N3 → N4 → activitate → servicii + controale
│   ├── report.py          # rapoartele Excel (plan / rezultate)
│   ├── sap_ids.py         # toate ID-urile SAP GUI (din Script1.vbs)
│   ├── sap_session.py     # conexiune, bara de stare, pop-up-uri, tabele SAP
│   ├── cj02.py            # acțiunile din tranzacție (creare WBS, rețea, servicii)
│   ├── runner.py          # execuția planului cu tratarea erorilor
│   └── cli.py, logging_setup.py, excel_recalc.py, textutil.py, model.py
└── tests/                 # pytest: deviz sintetic + SAP GUI simulat
```

Testele rulează fără SAP și fără Windows: `pip install pytest`, apoi `python -m pytest`.
