# moje-obligacje-data

Publiczne dane dla aplikacji **Moje obligacje** (Android): parametry i stopy procentowe wszystkich
niewykupionych serii (plus miesiąc po wykupie ostatniej transzy) detalicznych obligacji skarbowych (OTS, ROR, DOR, TOS, COI, EDO, ROS, ROD).

- `data/series.json` — seria → stopy kolejnych okresów, marża, opłata za wcześniejszy wykup, cena zamiany.
- `data/manifest.json` — wersja + sha256; aplikacja pobiera `series.json` tylko przy nowszej wersji.
- Publikacja: GitHub Pages, codziennie przez `.github/workflows/data.yml` (bez zmian treści nic się nie publikuje).

Źródła: arkusz MF „Dane dotyczące obligacji detalicznych” (gov.pl), listy emisyjne (obligacjeskarbowe.pl;
też nowa oferta, zanim trafi do arkusza) i tabele odsetkowe (obligacjeskarbowe.pl, kotwice kontrolne).
Szczegóły i walidacja: `tools/data/README.md`.

Dane mają charakter informacyjny; wiążące są listy emisyjne i rozliczenia agenta emisji.
