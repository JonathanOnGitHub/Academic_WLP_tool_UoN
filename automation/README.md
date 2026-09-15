# WLP Automator

Headed-Playwright driver for the Academic WLP web app
(https://jonathanongithub.github.io/Academic_WLP_tool_UoN/).

It walks through a configurable list of "straightforward" workload tabs,
uploads local data files into each one, clicks the relevant
**Analyse / Calculate** button, waits for the result to render, then
moves on. The browser stays open afterwards so you can inspect.

It does **not** modify the WLP web app, and it does **not** upload
anything to an external AI service — every action goes through a
real Chromium instance via Playwright's native APIs.

## Layout

```
automation/
├── .venv/                  # Python virtualenv (created by setup)
├── config/
│   └── tabs.yaml           # Tab → file mapping (edit this)
├── wlp_automator.py        # Main driver
├── run.sh                  # Convenience launcher
└── README.md
```

## One-time setup

The repo's Python is PEP 668-locked, so we use a venv inside
`automation/`:

```bash
cd automation
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install playwright pyyaml
playwright install chromium
```

## Running

```bash
./run.sh                                # use default config
./run.sh --config config/tabs.yaml      # explicit
./run.sh --dry-run                      # just print the plan
```

The first run downloads the Chromium binary (~170 MB) and stores it
under `~/.cache/ms-playwright/`.

## Editing the config

`config/tabs.yaml` is the only file you should normally touch. The
shape of each entry is:

```yaml
- key: teaching_timetable          # unique label, used in logs/screenshots
  enabled: true                    # set false to skip this tab
  main_tab: teaching               # data-maingroup of the top-level tab
  sub_tab: panel-teaching          # data-panel of the sub-tab (or '' if none)
  file_input: '#tlFileInput'       # <input type="file"> to upload into
  action: '#tlAnalyseBtn'          # Analyse/Calculate button (or '' if auto)
  wait_for: '#tlStatsBar'          # DOM element that signals result is ready
  files:                           # paths are resolved against data_dir
    - MPharm_Autumn_semester_as_of_010626.html
```

### Special kind: `citizenship_paste`

The School Citizenship tab's XLSX-upload path overwrites on each upload,
so to load all three categories (Teaching / Research / School) you have
to use the app's textarea-paste path instead. For that, set `kind:
citizenship_paste` and use a `paste:` map keyed by category:

```yaml
- key: citizenship_school_roles
  kind: citizenship_paste
  main_tab: citizenship
  sub_tab: panel-citizenship        # the 'School Citizenship' sub-tab
  action: '#citAnalyseBtn'
  paste:
    teaching: Teaching_roles.xlsx
    research: Research_roles.xlsx
    school:   School_and_Citizenship_roles.xlsx
```

The script reads each XLSX, converts it to TSV, drops it into the
matching `#cit-teaching` / `#cit-research` / `#cit-school` textarea, then
clicks the single Analyse button — which combines all three with
cross-category dedup.

### Special kind: `pgr_training_paste`

The PGR Training tab has three textareas (`#pgrtr-t1`, `#pgrtr-t2`,
`#pgrtr-t3`) that expect pasted tables. For that, set `kind:
pgr_training_paste` and provide a single XLSX with one sheet per
textarea:

```yaml
- key: citizenship_pgr_training
  kind: pgr_training_paste
  main_tab: citizenship
  sub_tab: panel-pgr_training
  action: '#pgrtrAnalyseBtn'
  files:
    - PGR_training_2025_26.xlsx    # sheets become t1, t2, t3 in order
```

Cells with non-breaking spaces, tabs, or newlines are stripped before
pasting so the app's parser can find the column headers cleanly.

### To swap in your own data

1. Edit `data_dir:` at the top, **or**
2. Change the `files:` (or `paste:`) list under each tab.

A tab whose files all exist will run; missing files fail that tab
loudly so you notice.

## How it waits for results

After clicking the action button, the script polls the `wait_for`
selector's text content until it is non-empty and not just `"—"`. This
covers every tab in V1 because each one populates a `stats-bar` element
once analysis finishes.

## Troubleshooting

* **Browser won't open on a headless server** — make sure X / Wayland is
  available. On a remote box, tunnel X11 or use VNC.
* **Element not found** — selectors were captured from the live app DOM
  in late 2026. If the upstream HTML changes, update the relevant
  `file_input` / `action` / `wait_for` lines.
* **File upload seems to do nothing** — confirm the filename's extension
  is on the input's `accept=` list (e.g. `.html` for the teaching tab).

## What's intentionally not in V1

Per the original brief and current data availability:

* **Combined Totals** merge — same V1 brief as before.

This can be added later — the per-tab structure makes it a copy-paste
addition to `tabs.yaml`.
