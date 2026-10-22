# tools/data — dane serii dla aplikacji

`build_series.py` buduje `data/series.json` (parametry i stopy każdej serii) oraz `data/manifest.json`.
Pliki leżą na GitHub Pages; telefon pobiera je sam, bez własnego backendu. Build działa bez nadzoru
(GitHub Actions codziennie); człowiek widzi tylko maile o nieudanym przebiegu, więc każda wątpliwość
w danych kończy się `exit 1` i niczego nie publikuje.

## Źródła

1. **Arkusz MF „Dane dotyczące obligacji detalicznych”** (BIFF `.xls`, ~1 MB), link na
   https://www.gov.pl/web/finanse/obligacje-detaliczne1 (sekcja „Materiały”). Skrypt szuka linku po tekście,
   bo UUID załącznika (`/attachment/…`) się zmienia. Arkusze `OTS, ROR, DOR, TOS, COI, EDO, ROS, ROD`:
   seria, początek sprzedaży, cena zamiany, stopy wszystkich ogłoszonych okresów (ułamki), marża.
   Sprawdzone 2026-10-22: plik istnieje (`Last-Modified: 2026-10-09`), więc `docs/spec-wzory.md` § 8 pkt 4
   („nie znalazłem tabeli XLS”) jest nieaktualne, a `docs/zrodla-danych.md` Q3 ma rację.
2. **Listy emisyjne (PDF)** z https://www.obligacjeskarbowe.pl/listy-emisyjne/ (+ archiwum): **opłata za
   wcześniejszy wykup**, której w arkuszu nie ma. Z LE bierzemy też stopę 1. okresu, marżę, cenę zamiany
   i miesiąc sprzedaży — do krzyżowej kontroli z arkuszem. Mapa seria → strona LE pochodzi z etykiet
   `<option>`, bo identyfikatory na stronie mają literówki (`edo07829` = EDO0829, `edo01028` = EDO1028).
3. **Tabele odsetkowe (PDF)** z https://www.obligacjeskarbowe.pl/tabela-odsetkowa/ — **kotwice kroczące**:
   w każdym przebiegu najnowsze ogłoszone stopy każdego typu porównywane z oficjalną tabelą
   („Oprocentowanie w bieżącym okresie”, „Okres odsetkowy”, kod serii). OTS nie ma tabel.

### Stopy z tabel odsetkowych (`"rateSources"`)

Arkusz MF ma stopy okresów zaczynających się 1. dnia miesiąca dopiero po swojej aktualizacji (~9.), a tabele
odsetkowe tych okresów są na stronie ok. 20. poprzedniego miesiąca (2026-10-22: 27 stóp okresów od 1.11,
których arkusz z 2026-10-09 nie ma). Dlatego dla serii w zakresie kolejne okresy po ostatniej stopie z arkusza
(ciągle, start najwyżej 45 dni naprzód) są brane z tabeli: „Oprocentowanie w bieżącym okresie”. Takie stopy
przechodzą tę samą walidację (zakres, spójność `stopa − marża` w grupie CPI/NBP, reguły opublikowanego
indeksu), a seria dostaje `"rateSources": ["XLS", …, "TO"]` (pole tylko, gdy któraś stopa jest z tabeli).
Arkusz wygrywa: gdy później poda ten sam okres, zmiana względem opublikowanej wartości to błąd A.
Kod serii i numer okresu bierzemy **z PDF-a**, nigdy z etykiety w liście rozwijanej (tabela 23869 jest tam
opisana jako DOR0728, a PDF to DOR0727) — niezgodna tabela jest traktowana jak brak tabeli. Stopy bywają
drukowane bez groszy („4%”). Pobierane są tylko tabele, których PDF-a nie ma jeszcze w cache (tabele są
niezmienne), więc codzienny przebieg pyta serwer o kilka nowych tabel w miesiącu. Opublikowana stopa z tabeli
nie znika, gdy tabela jest chwilowo niedostępna (zostaje z poprzedniej publikacji).

### Serie z samego LE (`"source": "LE"`)

Arkusz MF dostaje serie nowego miesiąca dopiero ok. 9., a listy emisyjne na następny miesiąc są ok. 21.–25.
poprzedniego. Serie bieżącego i następnego miesiąca, których arkusz jeszcze nie ma, są budowane z LE:
stopa 1. okresu, marża, opłata, cena zamiany, miesiąc sprzedaży (z tekstu „w dniach od 1 do 30 listopada
2026”). Dzięki temu kupujący 1. dnia miesiąca widzi wartości. Gdy arkusz je dostanie, zastępuje LE
(`"source": "XLS"`), a kontrola LE ↔ arkusz działa jak dla każdej serii.
Kontrole LE z samym sobą niczego nie dowodzą, więc stopa 1. okresu serii LE-only jest sprawdzana niezależnie:
z tabelą odsetkową okresu 1, gdy już jest (niezgodność = błąd A), i zawsze z ofertą tego typu z poprzedniego
miesiąca (skok > 1,5 pp = błąd A, np. odczyt 0,40% zamiast 4,00%). Bez tabeli i bez poprzedniej oferty
seria nie jest publikowana (alert B). Oferta **następnego** miesiąca (jeszcze nie w sprzedaży) jest
publikowana dopiero, gdy jej stopę potwierdzi tabela odsetkowa okresu 1 (tabele są na stronie ok. 20.,
przed miesiącem sprzedaży); sama kontrola skoku nie wystarcza. Do tego czasu tylko `WARNING`, bez alertu.
Komunikatu MF o ofercie nie używamy jako drugiego źródła (slug URL nieprzewidywalny, format niesprawdzony).

## Zakres

Wszystkie serie 8 typów, które nie są jeszcze w całości wykupione, **plus 1 miesiąc**: miesiąc z kodu serii
(`MMYY` = miesiąc wykupu ostatniej transzy) + 1 ≥ bieżący miesiąc. Zapas, bo wypłata ostatnich transz
przesuwa się na następny miesiąc (weekend), a użytkownik może jeszcze oglądać lot po wykupie.
TOS i OTS mają jedną stałą stopę; w JSON jest powtórzona dla każdego okresu (TOS: 3), bo silnik czyta
stopę per okres.

## Uruchomienie

```
pip install xlrd==2.0.2 pypdf==6.21.2   # jedyne zależności, czysty Python; przypięte (regexy pod pypdf)
python tools/data/build_series.py              # pełny build
python tools/data/build_series.py --offline    # tylko z cache
python tools/data/build_series.py --today 2026-11-10   # inny dzień odniesienia dla zakresu
python tools/data/build_series.py --alerts-out /tmp/alerts.out   # domyślnie tools/data/alerts.out
python -m unittest discover -s tools/data -p "test_*.py"
```

Testy: `test_build_series.py` (parsery i reguły na próbkach z `testdata/`; jedyne uruchamiane w CI),
`test_integration.py` (pełne buildy `--offline` na prawdziwym cache, np. alert „nieznany typ”) oraz
`test_audit.py`, `test_audit2.py`, `test_audit3.py` (audyty 2026-10-22; część testów buduje `--offline` na prawdziwym cache,
bez cache są pomijane; testy poziomu workflow W1–W3 są pominięte — pilnuje ich `data.yml`). Żaden test nie
zależy od zegara ani od treści opublikowanego `data/series.json`, gdy oczekiwanie zależy od danych:
daty są wstrzykiwane (`--today` wyliczane z daty arkusza w cache), oferty porównawcze podawane jako fixture.

Cache: `tools/data/cache/` (w `.gitignore`). Strona MF i arkusz odświeżane, gdy cache starszy niż 12 h;
bieżący listing LE i listing tabel odsetkowych co 20 h; archiwum LE raz w tygodniu. Strony HTML listów
i tabel oraz ich PDF-y są niezmienne: poprawnie zwalidowane zostają w cache na zawsze, ponownie pobierane
są tylko te, które nie przeszły walidacji. Pierwszy build to ~1000
żądań (1 s przerwy, prawdziwy User-Agent, ~25 min); codzienne przebiegi: kilka-kilkanaście żądań.

**Odporność cache:** pobrany plik musi przejść kontrolę swojego rodzaju, zanim trafi do cache: `%PDF`,
`D0CF11E0` dla XLS, a strony HTML dodatkowo znacznik treści (strona MF: „obligacji detalicznych”, listingi:
opcje `<select>`, strona listu/tabeli: link do PDF). Strona przerwy technicznej albo WAF z kodem 200 nie jest
więc zapisywana ani ponawiana w tym przebiegu; następny przebieg pyta ponownie. Zapis przez plik
tymczasowy + rename; plik w cache, który nie przechodzi kontroli, jest ignorowany; `.txt` z PDF-a powstaje
tylko z niepustego tekstu. Gdy sieć zawiedzie, używana jest ostatnia poprawna kopia z cache (alert B).
Każdy nieprzewidziany wyjątek kończy się `FATAL: …` i `exit 1`, nie tracebackiem.

## Dwie klasy problemów

- **(A) złe dane** → `ERROR:` + `::error::…` (powód widać w mailu o nieudanym przebiegu), `exit 1`,
  nic nie jest zapisywane ani publikowane.
- **(B) nieaktualne / niezweryfikowane / nie udało się pobrać** → `ALERT:` + `::warning::…`, wpis w
  `tools/data/alerts.out` (jedna linia na alert, nadpisywany w każdym przebiegu, w `.gitignore`)
  i w `$GITHUB_STEP_SUMMARY`. Wstrzymane są **tylko dotknięte serie albo stopy** (brak = „n/a” w aplikacji,
  nigdy zgadywanie), reszta jest publikowana, `exit 0`. Workflow zamienia niepusty `alerts.out` w osobny,
  nieudany job „alerts” (mail), który nie blokuje publikacji dobrych danych.

### (A) — exit 1

- luka w stopach (pusty okres przed wypełnionym);
- liczba znanych stóp ≠ liczba okresów rozpoczętych dla zakupu 1. dnia sprzedaży na dzień `Last-Modified`
  arkusza (dopuszczalna +1, gdy MF wpisał stopę z wyprzedzeniem);
- stopa poza [0; 0,25]; brak marży (ROR/DOR/COI/EDO/ROS/ROD);
- kod serii niezgodny z miesiącem sprzedaży + czas życia typu;
- **spójność indeksu**: dla każdego miesiąca startu okresu (od 2. okresu) `stopa − marża` musi być identyczna
  we wszystkich seriach grupy CPI (COI/EDO/ROS/ROD) i osobno grupy NBP (ROR/DOR) — to ten sam CPI albo ta
  sama stopa referencyjna. Dodatkowo indeks nie może się zmienić dla miesięcy już opublikowanych
  w poprzednim `data/series.json`. Stan 2026-10-22: 0 niespójności (CPI 109 miesięcy, NBP 24);
- **LE ↔ arkusz**: LE musi wymieniać kod swojej serii (inaczej to zły dokument), zgadzać się stopą 1. okresu,
  ceną zamiany, marżą i miesiącem sprzedaży; serie LE-only: patrz wyżej (tabela okresu 1, skok > 1,5 pp);
- **tabela odsetkowa się nie zgadza** z danymi (kotwice kroczące, głębokie, LE-only);
- **kotwice stałe** (`ANCHORS`, wartości z testów silnika i tabel odsetkowych): EDO1026, COI1026, ROS1026,
  TOS1026, EDO1036, TOS1029, OTS0127 oraz długożyjące ROR0127, DOR0127, COI1027, EDO1027, ROS1027, ROD1028
  (sprawdzone 2026-10-22 z 57 tabelami odsetkowymi, 0 rozbieżności). `rates` to prefiks. Kotwica serii poza
  zakresem albo wstrzymanej w tym przebiegu jest pomijana; seria kotwicy, której nie ma w danych z innego
  powodu, to błąd. Liczba pozostałych kotwic stałych nie ma znaczenia (żadnej bomby z datą) — ciągłość
  dają kotwice głębokie;
- liczba serii spadła o >5% względem poprzedniego `data/series.json`;

### (B) — alert, wstrzymanie tylko dotkniętych danych

- **nowy miesiąc indeksu niezweryfikowany**: każdy miesiąc indeksu (grupa CPI, grupa NBP), którego nie ma
  jeszcze w opublikowanym indeksie, musi mieć ≥ 1 stopę potwierdzoną tabelą odsetkową (kandydaci: serie
  zaczynające w tym miesiącu okres ≥ 2, najnowsza sprzedaż pierwsza; brak tabeli na stronie = następny
  kandydat). Bez tego wstrzymane są nowe stopy od najwcześniejszego niezweryfikowanego miesiąca wzwyż.
  **Miesiące już opublikowane nie są nigdy wstrzymywane** (są zablokowane: zmiana wartości = błąd A), a
  miesiąc wstrzymany wczoraj nie przejdzie dziś bez sprawdzenia tylko dlatego, że pojawił się nowszy.
  Pierwsza publikacja (brak `data/series.json`): weryfikowany tylko najnowszy miesiąc.
  Niezależnie od tego kotwice kroczące: dla ROR, DOR, TOS, COI, EDO, ROS, ROD po 2 serie z najnowszym
  ogłoszonym okresem (niezgodność = A, brak tabeli dla najnowszego miesiąca typu = `WARNING`);
- **kotwice głębokie (samoodnawialne)**: dla każdego typu z tabelami najstarsza seria w zakresie, wszystkie
  jej okresy z tabelami odsetkowymi (PDF-y w cache na zawsze, więc tanio). Gdy seria wypada z zakresu,
  następna najstarsza przejmuje rolę. Niezgodność = (A); żaden okres niezweryfikowany = alert. Tabele
  w starym układzie (sprzed ~2020) nie parsują się i są pomijane;
- list emisyjny serii niedostępny albo nieczytelny → seria już opublikowana: warunki LE z poprzedniej
  publikacji (nie zmieniają się) i dalej kontrola z arkuszem; seria nowa: wstrzymana;
- **opublikowana seria nigdy nie znika**, dopóki jest w zakresie: gdy z jakiegokolwiek powodu jej dziś
  nie ma (arkusz, listing LE), poprzedni wpis jest przenoszony bez zmian i jest alert;
- seria LE-only nieweryfikowalna (bez tabeli i bez poprzedniej oferty) → niepublikowana;
- **świeżość** (dostosowana do codziennego crona): `Last-Modified` arkusza starszy niż 35 dni; najnowsza
  seria w arkuszu starsza niż poprzedni miesiąc; od 3. dnia miesiąca brak jakiejkolwiek serii (arkusz
  albo LE) na bieżący miesiąc (dni 1–2: ostrzeżenie). W dniach 1–8 nic z tego nie strzela przy normalnym
  rytmie MF (arkusz z ~9. poprzedniego miesiąca ma ≤ 30 dni, oferta bieżąca jest z LE). Stary arkusz to
  dane stare, nie złe: brakujące nowe okresy są po prostu nieobecne;
- pobranie nie powiodło się i użyto starej kopii z cache.
- **nieznany typ obligacji**: seria w zakresie o prefiksie spoza 8 typów silnika — w dowolnym arkuszu XLS
  spoza tych 8 (także w nowym arkuszu) albo w listingu listów emisyjnych. Aplikacja i tak nie umie wycenić
  nieznanego typu, więc jego serie nie trafiają do `series.json`, a reszta jest publikowana. Alert w każdym
  przebiegu, dopóki problem trwa: `UNKNOWN BOND TYPE <prefiks>: <serie> — app needs a code update (BondType.kt)`.

### Znane typy historyczne (ignorowane, wszystkie serie wykupione)

W arkuszu MF: `KOS` (krótkoterminowe, do 2016), `POS` (premiowe, do 2023), `IR`, `RS` (lata 90.),
`DOS` (2-letnie stałe, do 05.2024), `TOZ` (3-letnie zmienne, do 07.2025), `TZ` (do 2015), `SP` (do 2010).
Kody z lat 90. (`IR0995`) są czytane jako 19YY. Jeśli któryś z tych prefiksów wróci z nową serią, alarm
„nieznany typ” zadziała tak samo jak dla zupełnie nowego.

### Parsowanie LE — pułapki

pypdf w części PDF-ów wstawia spację przed polskimi znakami (`nie wy ższ ą ni ż 2,00 zł`) — tekst jest
normalizowany przed dopasowaniem. Starsze LE ROS/ROD podają opłatę bez groszy (`nie wyższą niż 3 zł`),
EDO0134 ma literówkę `wynos 6,90%`. Opłata musi wystąpić w LE dokładnie jedną kwotą (COI/ROR/DOR mają dwie
klauzule, kwota musi być ta sama), inaczej błąd.

## Wersja i manifest (wspólny wzorzec aplikacji)

`series.json` ma na górze `"schema": 1` i `"version"` (liczba `YYYYMMDDNN`, rosnąca; gdy poprzednia
wersja jest datowana „w przyszłości”, po prostu +1). Każda seria ma `"source": "XLS" | "LE"`.
`manifest.json`: `{"version": …, "schema": 1, "files": {"series.json": {"sha256": "<hex>", "bytes": n}}}`.
Telefon czyta najpierw manifest (tani), pobiera `series.json` tylko przy nowszej wersji, sprawdza sha256
i podmienia plik atomowo.
Wersja **nie rośnie**, gdy treść się nie zmieniła: porównanie z poprzednim `series.json` z pominięciem
`generatedAt`, `version` i górnego `source` (nowy arkusz MF z identycznymi liczbami nie wymusza pobrania).
sha256 liczony jest zawsze z **kanonicznych bajtów** (UTF-8, LF), tych samych, które są publikowane; plik
na dysku z innymi końcami linii (checkout Windows z `core.autocrlf`) jest przepisywany do postaci
kanonicznej. `.gitattributes` w obu repo: `data/*.json -text` (git nie rusza końców linii).

## Kiedy uruchamiać

- ~9. dnia MF aktualizuje arkusz (jedyny znany punkt: 2026-10-09 — **harmonogram niepotwierdzony**;
  po kilku miesiącach sprawdzić `Last-Modified` w `cache/mf.xls.meta.json` i ewentualnie próg 35 dni);
- ~15. GUS publikuje finalne CPI (podstawa stóp COI/EDO/ROS/ROD na okresy od następnego miesiąca);
- ~21.–25. MF publikuje listy emisyjne i komunikat o ofercie na następny miesiąc.

Cron: **codziennie** `17 6 * * *` (`.github/workflows/data.yml` w repo `moje-obligacje-data`). Bez zmian
treści nic się nie commituje ani nie publikuje. Nowa oferta wchodzi z LE w dniu publikacji listów, stopy
nowych okresów z arkuszem MF. Przebieg nieudany = mail do właściciela repo; aplikacje działają na ostatnich
opublikowanych (poprawnych) danych.

W CI przed buildem działają tylko testy parsera (`test_build_series.py`); testy audytowe nie blokują
buildu. Alerty (B) dają osobny nieudany job "alerts", błędy (A) nieudany build.
