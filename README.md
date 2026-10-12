# Tankpot fuel prices

Small, public fuel price files used by the Tankpot app. No app code lives here.

- `eu/<CC>.json`: weekly average pump prices (with taxes, EUR per litre) for each EU country: Euro 95, diesel and LPG where reported. About three years of history; newest week last.
- `eu/index.json`: the countries available and their newest week.

A GitHub Action (`.github/workflows/update.yml`) runs daily, downloads the European Commission's Weekly Oil Bulletin and rewrites these files only when there's something new. The app downloads just the one country file it needs (a few KB) and keeps it on the phone for 12 hours.

Source: European Commission, DG Energy, [Weekly Oil Bulletin](https://energy.ec.europa.eu/data-and-analysis/weekly-oil-bulletin_en). Reuse of Commission data is allowed with attribution (Commission Decision 2011/833/EU).
