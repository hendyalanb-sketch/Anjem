# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**SiMADYA** (Sistem Manajemen Antar-Jemput RS Adhyaksa Jawa Timur) is a static single-page analytics dashboard for planning hospital employee shuttle routes. It visualizes survey data from 110 employees of RS Adhyaksa Jawa Timur (Mojokerto, East Java) who reported their home locations and pickup point preferences.

## Regenerating the Dataset

The only "build" step is regenerating `data.js` from the Excel survey file:

```bash
pip install openpyxl
python gen_data.py > data.js
```

`gen_data.py` reads `Survei Kebutuhan Rute dan Titik Jemput Pegawai RS Adhyaksa Jatim (Jawaban).xlsx`, geocodes each free-text pickup address via regex landmark matching (`LANDMARKS` list), and outputs four JS globals: `HOSPITAL`, `ROUTES`, `POINTS`, `RESPONSES`.

There are no tests, no linter, and no package manager — this is a zero-dependency static app.

## Architecture

Everything lives in two files:

- **`index.html`** — ~1500-line monolithic file containing all HTML structure, inline CSS (`<style>`), and inline JavaScript (`<script>`). No external JS files except `data.js`.
- **`data.js`** — auto-generated, never hand-edited. Contains the four globals consumed by `index.html`.

External CDN dependencies (no local install needed):
- Tailwind CSS (CDN, config inlined in `<script>`)
- Leaflet.js 1.9.4 + MarkerCluster 1.5.3 for maps

### Data model

| Global | Description |
|---|---|
| `HOSPITAL` | Single object: hospital name, lat/lng, address |
| `ROUTES` | Dict keyed by route ID (`mjk-kota`, `trowulan`, `mojoagung`, `mojosari`, `puri`) → `{name, color}` |
| `POINTS` | Array of aggregated pickup stops: `{label, lat, lng, area, route, count, members[]}` |
| `RESPONSES` | Array of 110 individual survey records with compact field names (`n`=name, `u`=unit, `lat`/`lng`, `route`, etc.) |

### UI architecture

Navigation is tab-based. `setTab(tab)` toggles `.active` on `.nav-link[data-tab]` and `[data-panel]` elements. The seven panels are:

| `data-tab` | Panel content |
|---|---|
| `overview` | KPI cards + mini-map + bar charts |
| `map` | Full Leaflet map with filter sidebar |
| `routes` | Route recommendation cards (`renderRoutes()`) |
| `directions` | One-way route polyline map (`renderOneWayRoutes()`) |
| `bus-route` | Bus stop sequence map + bottom sheet (`renderBusRoute()`) |
| `data` | Searchable employee table (`renderDataTable()`) |
| `insights` | Survey analytics charts (`renderInsights()`) |

Global state is in the `STATE` object: holds Leaflet map instances (`map`, `miniMap`, `oneWayMap`, `busRouteMap`), active route filters (`routes: Set`), and layer visibility flags.

### Adding or editing pickup points / route assignments

All geocoding logic and landmark coordinates live in `gen_data.py`:
- `LANDMARKS` — ordered list of `(regex, label, lat, lng, area, route_id)` tuples; first match wins
- `ROUTES` — dict of route IDs to display name and color

After editing `gen_data.py`, regenerate `data.js` and refresh `index.html` in the browser.

## Development

Open `index.html` directly in a browser (no server required for local viewing). For development with live reload, any static file server works:

```bash
python -m http.server 8000
```

The GitHub Actions workflow (`.github/workflows/blank.yml`) is a placeholder stub — there is no automated deployment pipeline.
