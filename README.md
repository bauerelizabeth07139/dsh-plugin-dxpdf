# dsh-plugin-dxpdf

A DeepSeek Harness (DSH) host plugin that exposes one tool, **`dxpdf_convert`**,
which converts a Microsoft Word `.docx` file to `.pdf` with the third-party
[dxpdf](https://github.com/nerdy-pro/dxpdf) engine — a compiled Rust + Skia
Python extension that needs no Microsoft Office, no LibreOffice and no network.
It also repairs the schema-invalid constructs WPS Office writes, and gives back
the blank pages Word inserts at odd/even section breaks.

Engine: [github.com/nerdy-pro/dxpdf](https://github.com/nerdy-pro/dxpdf) ·
PyPI: [`dxpdf`](https://pypi.org/project/dxpdf/) ·
License: [MIT](LICENSE)

## Requirements

- **Node.js ≥ 22** — the Harness that hosts the plugin.
- **Python ≥ 3.8 on `PATH`** — or pointed at explicitly with `pythonPath` or
  `DXPDF_PYTHON`. Any interpreter will do, as long as it can import the engine.
- **The dxpdf engine installed into that interpreter**:

  ```console
  python -m pip install dxpdf
  ```

- **`lxml`, for the Word-parity pagination pass**:

  ```console
  python -m pip install lxml
  ```

  `py/paginate.py` imports `lxml`. Without it the pagination pass is skipped:
  the check fails with `ModuleNotFoundError: No module named 'lxml'`, the
  conversion still completes, and the result carries `paginationSkipped`
  explaining why. The pagination pass is the only part of the plugin that needs
  `lxml`; conversion, the CLI shim and the schema repair do not.

## Install

The plugin is a DSH bundle. Install it into the profile you run, switch the
bundle on, then install the engine into the interpreter the plugin will use.

### Desktop app

**Plugins → Add plugin** → `https://github.com/bauerelizabeth07139/dsh-plugin-dxpdf`,
then switch the new bundle on in the same list. The Desktop app boots the
reserved `desktop` profile, so that is where the bundle has to be enabled.

### CLI

```console
dsh plugin --profile web add bauerelizabeth07139/dsh-plugin-dxpdf
```

On a machine with no `git`, that shorthand cannot resolve: pnpm runs
`git ls-remote` and fails with `'git' is not recognized`. Install from the
GitHub tarball instead:

```console
dsh plugin --profile web add https://codeload.github.com/bauerelizabeth07139/dsh-plugin-dxpdf/tar.gz/main
```

Replace `main` with a commit SHA to pin an exact revision.

### Install the engine

The plugin shells out to the dxpdf engine, which is a Python package. Install
it — with `lxml` — into the interpreter the plugin will use:

```console
python -m pip install dxpdf lxml
```

### Uninstall

```console
dsh plugin --profile web remove dsh-plugin-dxpdf
```

## Why this plugin exists

The `dxpdf` wheel on PyPI ships only the compiled extension module
(`dxpdf/dxpdf.pyd`) and the `convert`/`convert_file` functions. It declares no
console script, so `pip install dxpdf` leaves the documented `dxpdf` command
unavailable. This plugin bridges that gap twice over:

- the package ships its own copy of the CLI shim (`py/dxpdf_cli.py`), which the
  plugin runs with any Python it can find, so `python -m pip install dxpdf` is
  the only preparation a conversion needs;
- the plugin drives the engine with an explicit argv, so a model can convert
  documents without quoting hazards — only the discovered `dxpdf.cmd`/`.bat`/
  `.exe` path goes through `cmd.exe`, and that path double-quotes every token.

It also repairs the documents dxpdf would otherwise refuse — see
[WPS Office compatibility](#wps-office-compatibility).

## WPS Office compatibility

dxpdf parses OOXML with strict serde schemas. Word accepts four constructs that
WPS Office writes and dxpdf rejects, so a chapter that opens perfectly in an
editor fails to convert:

| Construct | What WPS emits | What dxpdf says |
|---|---|---|
| Repeated property element | `<w:rPr>` written with two `<w:sz>` children (`21` then `20`), the second after `<w:b/>` and `<w:color/>` | ``failed to deserialize XML: duplicate field `sz` `` |
| Property element not first | a TOC entry's `<w:pPr>` after its three field-char runs | ``duplicate field `$value` `` |
| Run nested in a run | `<w:r><w:r>…</w:r><w:r><m:oMath>…</m:oMath></w:r></w:r>` wrapped around a text-and-equation sequence | ``unknown variant `r`, expected one of `t`, `delText`, …`` |
| Equation wrapped in a run | `<w:p><w:r><m:oMath>…</m:oMath></w:r></w:p>`, when OMML math belongs to the paragraph | ``unknown variant `oMath`, expected one of `t`, `delText`, …`` |

All four are systematic — a whole textbook series can carry them, in
combination. The circuit chapter needed 206 repairs across three of the four.

`py/normalize.py` rewrites only the offending elements and copies every other
ZIP entry through byte for byte. A repeated property is removed (the first
occurrence wins by default); a run-wrapped equation is **unwrapped** to its
paragraph; a nested run is **spliced** into its parent, handing its `w:rPr`
down to child runs that lack one; a misplaced `w:pPr` is **moved to the front**.
The repaired copy is converted in a temporary directory and the source file is
never modified.

### Why `$value` names no element

quick-xml's serde binds an element's character data to the reserved field
`$value`. A `w:pPr` sitting mid-paragraph interrupts that data, so dxpdf sees
the paragraph's flattened content twice and reports ``duplicate field
`$value` `` — an error that names a serde pseudo-field, not an element, and
therefore cannot be repaired by name. It is fixed by the order-only reorder
rule, which runs on every pass precisely because no name is available to target.

### Rules that need no permission, and rules that do

Repairs 2–4 are structure-preserving: unwrapping a run, splicing a nested run
and moving a schema-fixed property to the front cannot change what a document
renders, so they run on every pass. Only repair 1 *deletes* something, so only
repair 1 waits for the engine to name the field it rejected.

### Which duplicate wins

For the duplicates seen in the wild, `Heading1` carried `32/32` and `Heading2`
`26/26` — identical values, which shows the repeat is a redundant re-emission
rather than a deliberate override. The first occurrence is also the one in
canonical schema position, paired with `w:szCs`. The default is therefore
`first`; the measured difference either way is at most half a point, and
`duplicatePolicy: last` flips it.

### Nothing is removed on a guess

The engine names the field it choked on. The plugin feeds that name back —
``duplicate field `sz` `` becomes `--only sz` — and repairs *only* what was
actually rejected, adding one name per round. A document is re-tried after each
round, and a rejected document fails during parsing, before any layout or
painting, so extra rounds cost milliseconds.

This matters: sweeping a built-in list of "should appear once" elements instead
is how a document gets quietly damaged. `w:tblStylePr` (one per conditional
table region) and `w:rsid` (the whole revision-id history) are repeating
elements that a naive list marks as duplicates. The targeted path never touches
them, and the bulk fallback excludes them explicitly.

### One pass is not always enough

The rules feed each other. Unwrapping a nested run exposes equations that then
sit at paragraph level and need unwrapping themselves; moving a property can
reveal the next one out of order. Each pass re-parses what the previous one
produced and stops when a pass finds nothing, so the result is a fixed point
rather than one rule's optimistic guess.

The duplicate `w:pPr` blocks these documents also carry are schema-invalid but
dxpdf tolerates them, so a targeted repair correctly leaves them alone.

## Word-compatible pagination

Word and dxpdf disagree in one place, and it shows up as a page that should be
there and is not. A section break whose `w:sectPr/w:type` is `oddPage` or
`evenPage` promises the *next* section a page of that parity; Word inserts a
blank page when the section before it ended on the wrong one. dxpdf documents
both as *treated as nextPage*, so it breaks to the very next page and the blank
page never appears. The circuit chapter carries eleven such breaks.

dxpdf is a compiled wheel, so the fix is not to change its layout but to hand it
a document whose plain `nextPage` breaks already land on the required parity.
Two measured facts make that exact rather than a search:

- an empty paragraph carrying `w:pageBreakBefore` reliably costs exactly one
  page, so one filler per offending boundary reproduces Word's pagination;
- adding a filler shifts every later section by exactly one page and flips its
  parity, so a boundary never needs a second filler.

`py/paginate.py` measures the page each section ends on by converting a prefix
of the document that stops at that boundary, then walks the boundaries once and
inserts a filler wherever the next section would otherwise start on the wrong
parity. Each probe is taken with the fillers decided so far already in place, so
the shift is observed rather than modelled.

| construct | dxpdf | this fix |
|---|---|---|
| `w:br w:type="page"` | honoured — verified at paragraph end and paragraph start | untouched |
| `w:sectPr w:type="nextPage"` | honoured | untouched |
| `w:sectPr w:type="oddPage"` | treated as `nextPage` | parity filler where needed |
| `w:sectPr w:type="evenPage"` | treated as `nextPage` | parity filler where needed |

### How it is wired in

`dxpdf_convert` runs the pass itself, between the schema repair and the final
conversion, and reports what it did:

```
Converted … (12.2 MiB, 269929 ms, image_dpi=220, 420 pages).
Repaired 58 schema violation(s) in a copy first; the source file was not modified.
Restored 6 blank page(s) so the 10 odd/even section break(s) land on the page parity Word promises.
```

It is gated so the common document pays almost nothing. `paginate.py --check`
answers from a **byte scan of `document.xml`** for `oddPage`/`evenPage`; a
document with neither never reaches the measurement pass, and the tool reports
`paginated: false, boundaries: 0`. Measured on the corpus:

| document | parity breaks | pages | cost |
|---|---|---|---|
| `核探测仪器-数学篇.docx` | 0 | 458 | 25 s total, no measurement pass |
| `核探测仪器-电路篇.docx` | 10 | 414 → 420 | 270 s total, 12 probes |

Two guards bound the cost when the pass does run, and they are coupled: the
timeout must exceed one engine run per parity break, or a document that is
inside the section cap fails on the clock instead of converging.
`maxPaginateSections` (default `48`) refuses a document with more parity breaks
than that rather than spending an unbounded number of engine runs, and the tool
then reports `paginationSkipped` explaining the refusal instead of silently
under-paginating. `paginateTimeoutMs` (default 45 minutes) bounds the pass
itself. At roughly 20 s a probe, 48 breaks need about 16 minutes, so raising the
cap much further means raising the timeout with it.
`matchWordPagination: false` turns it off entirely.

The steps can also be run by hand, which is what the tests do:

```powershell
python py\normalize.py in.docx --output stage1.docx --quiet
python py\paginate.py  stage1.docx --output final.docx --report report.json
python py\paginate.py  in.docx --check        # boundaries only, no conversion
```

The report names every boundary, the page its section ends on, the page the next
section starts on, and whether a filler was inserted. `--python` and `--cli`
override how the engine is invoked, matching the plugin's own runner settings.

## The tool

### `dxpdf_convert`

| Parameter | Type | Required | Meaning |
|---|---|---|---|
| `input` | string | yes | Source `.docx`. Absolute, or relative to the session working directory. |
| `output` | string | no | Destination `.pdf`. Defaults to the input path with a `.pdf` extension. An existing file is overwritten. |
| `image_dpi` | number | no | Resolution embedded raster images are downsampled to, 1–2400. Defaults to `220`, which is what Word uses. Raise it (e.g. `300`) for print quality, lower it (e.g. `96`) for smaller files. Images are never upsampled past their source resolution. |

The tool returns the resolved absolute `input` and `output`, the `imageDpi`
used, the produced `bytes`, `elapsedMs`, an optional `pages` count read back
from the PDF, and `runner` — which records how the engine was located, for
troubleshooting a machine's setup.

When the document had to be repaired first, `normalized` is `true` and
`repairs` counts the schema violations fixed. When its odd/even section breaks
needed blank pages, `paginated` is `true` with `fillers` and `boundaries`; if
the pass was refused or failed, `paginationSkipped` says why. The rendered
summary states both, so neither is silent.

Failures are ordinary tool errors, not turn-ending ones: a missing input, a
non-DOCX input, an out-of-range `image_dpi`, a non-zero exit from the engine,
and a timeout all surface as messages the model can read and act on. The
conversion observes the caller's cancellation signal.

## Runner resolution order

The first candidate that exists on disk wins. `pythonPath`, `cliPath` and
`executable` come from configuration; the rest from the environment.

| # | Candidate | How it is spawned |
|---|---|---|
| 1 | configured `pythonPath` + `cliPath` | `spawn(python, [cli, …])`, direct — no shell |
| 2 | `DXPDF_PYTHON` + `DXPDF_CLI` | `spawn(python, [cli, …])`, direct — no shell |
| 3 | configured `executable`, else `DXPDF_EXE`, else `dxpdf.cmd` / `dxpdf.bat` / `dxpdf.exe` found on `PATH` | `cmd.exe /d /s /c` with `windowsVerbatimArguments`; every token double-quoted |
| 4 | any Python on `PATH` (or `pythonPath` / `DXPDF_PYTHON`) + **the CLI shim this package ships** (`py/dxpdf_cli.py`) | `spawn(python, [py/dxpdf_cli.py, …])`, direct — no shell |
| 5 | a Python on `PATH` + `dxpdf_cli.py` on `PATH` | `spawn(python, [cli, …])`, direct — no shell |

Candidates 1 and 2 need both files to exist; 3 needs the executable to exist;
4 needs a resolvable interpreter and the bundled shim, which ships with the
package. Row 4 is what makes `pip install dxpdf` alone sufficient: the wheel
declares no console script, so there is no `dxpdf` command to find, and the
plugin supplies the missing front end. Rows 1, 2, 4 and 5 pass an argv array
straight to `spawn`, so no shell ever parses a filename; row 3 is the only one
that involves `cmd.exe`.

When nothing resolves, the tool fails with the candidate list it tried:

```
dxpdf is not installed where this plugin can find it. Tried configured
pythonPath/cliPath: …; DXPDF_PYTHON/DXPDF_CLI machine variables: …; …
Install it with "python -m pip install dxpdf", or point the plugin at it with
the pythonPath/cliPath options.
```

## Configuration

Every field is optional; the row in `cordis.patch.yml` needs no `config` block.
To override a default, add one:

```yaml
- insert:
    - id: dsh-plugin-dxpdf
      name: 'dsh-plugin-dxpdf'
      config:
        timeoutMs: 300000
        pythonPath: 'C:\Python312\python.exe'
        cliPath: 'C:\Python312\Scripts\dxpdf_cli.py'
        duplicatePolicy: first
        announcePolicy: true
```

| Field | Default | Meaning |
|---|---|---|
| `timeoutMs` | `120000` | Wall-clock budget for one conversion. |
| `paginateTimeoutMs` | `2700000` | Budget for the pagination pass, which runs the engine once per parity section. Must exceed roughly 20 s × the section count. |
| `maxPaginateSections` | `48` | Refuse the pagination pass beyond this many parity breaks, reporting `paginationSkipped` instead. |
| `matchWordPagination` | `true` | Converge odd/even section breaks onto Word's pagination. Documents with no such break are detected by a byte scan and cost nothing. |
| `duplicatePolicy` | `first` | Which occurrence of a repeated property the repair keeps: `first` or `last`. |
| `announcePolicy` | `true` | Whether to register the standing DOCX-to-PDF system-prompt section. |
| `pythonPath` | unset | Explicit interpreter, tried first. |
| `cliPath` | unset | Explicit CLI script, tried first (paired with `pythonPath`). |
| `executable` | unset | Explicit `dxpdf` shim, tried in the `cmd.exe` row. |

The two booleans are disabled only by `false`; `duplicatePolicy` is `first`
unless it is exactly `last`; a non-positive or non-numeric timeout or cap falls
back to its default. `announcePolicy: false` suppresses the standing
DOCX→PDF system-prompt section — `dxpdf:docx-to-pdf-policy`, order `1250` —
that makes this plugin the deployment's converter.

## Priority: the deployment's DOCX-to-PDF converter

The Harness has no numeric priority for tools. What decides which converter a
model reaches for is what the model reads, so this plugin claims the job in the
two places that actually matter.

**1. A standing system-prompt section.** A tool description only competes with
other tools; it does not settle a choice against a *skill*. The bundled
`office-docx` skill tells the model to render a PDF deliverable with bundled
LibreOffice (`<cli> convert --input report.docx --output report.pdf`), and a
skill's instruction would otherwise win. The plugin therefore registers an
always-present prompt section, `dxpdf:docx-to-pdf-policy`, at order `1250` —
between the write tool's section (`1200`) and the edit tool's (`1300`):

```
## DOCX to PDF

Convert every DOCX-to-PDF request with the `dxpdf_convert` tool. …
Do not use bundled LibreOffice, `soffice`, or a Python rendering library to
turn a .docx into a .pdf. …
```

It is scoped deliberately. The `office-docx` skill still governs creating,
editing and structurally checking Word documents; XLSX and PPTX still convert
through LibreOffice; and the policy names its own escape hatch — a document the
engine cannot handle (charts, SmartArt, tracked changes) falls back to the
skill's bundled-LibreOffice path, with the model told to say so.

**2. A directive tool description.** `dxpdf_convert`'s description ends by
stating that it is the deployment's converter and should be preferred over
LibreOffice or a Python rendering library for any `.docx` → `.pdf`.

The section is registered with `ctx.systemPrompt.section(...)`, which returns
its own Cordis effect disposer, so it disappears with the plugin rather than
accumulating across reloads. `announcePolicy: false` turns it off.

## Tests

The suites are plain Node scripts with no test framework and no dependencies of
their own. They convert real DOCX fixtures with the real engine, so the engine
must be installed before they will pass:

```powershell
npm test              # every suite below, in order

# Conversion core — 12 checks: runner discovery from DXPDF_* variables and from
# PATH, the .cmd shim strategy, both spawn strategies, page counting, image_dpi
# validation, error paths, overwrite behaviour.
npm run test:core

# Tool contract — 13 checks: loads lib/index.js through a resolve hook that
# substitutes a stub for @deepseek-ai/dsh-tools, captures the registered
# definition and the system-prompt policy, and runs its execute/render/presentCall.
npm run test:tool

# Document repair — 9 checks: the raw engine rejects wps-defects.docx, the
# plugin converts it anyway, the source is left untouched, both policies work,
# and a repaired copy needs no second repair.
npm run test:repair

# Word-parity pagination — 6 checks: the cheap check finds no break in an
# ordinary document, finds the odd-page break in parity-sections.docx, and the
# conversion restores exactly the blank page Word inserts (2 -> 3 pages).
npm run test:paginate

# Bundle composition — 10 checks over the installed profile wiring. Needs a
# profile that has this bundle installed and a copy of js-yaml.
npm run test:bundle
```

The engine location used by the suites is taken from the test overrides
`DXPDF_TEST_PYTHON`, `DXPDF_TEST_CLI` and `DXPDF_TEST_SHIM`; each falls back to
the machine variables the plugin itself reads (`DXPDF_PYTHON`, `DXPDF_CLI`,
`DXPDF_EXE`) and then to the shim this package ships. `test/convert.test.mjs`
reports the `.cmd` shim check as skipped when no shim is configured.

`test/bundle.test.mjs` inspects an installed profile, so it needs `DSH_PROFILE`
pointed at one (it defaults to `~/.dsh/profiles/desktop`, or
`DSH_PROFILE_NAME`), and a copy of `js-yaml`, which ships with the Harness
rather than with this plugin (`DSH_JS_YAML`):

```powershell
$env:DSH_PROFILE = '<profile directory>'
$env:DSH_JS_YAML = '<profile>\node_modules\js-yaml\dist\js-yaml.mjs'
npm run test:bundle
```

When either is missing it prints SKIP and exits 0, so `npm test` works on a
fresh clone.

`test/fixtures/wps-defects.docx` is generated by
`test/fixtures/make-wps-defects.py`, which injects one of each defect into
`single-page.docx`; `test/fixtures/make-parity-sections.py` builds
`parity-sections.docx`. Both generators use `lxml`. Regenerate a fixture after
changing its base document. CI runs `npm test` on Node 22 and Python 3.12 on
Ubuntu and Windows.

## Troubleshooting

### `dxpdf is not installed where this plugin can find it`

Every candidate in [Runner resolution order](#runner-resolution-order) was
tried and none existed on disk. Install the engine into the interpreter the
plugin resolves, or point the plugin at an interpreter and a CLI directly:

```yaml
config:
  pythonPath: 'C:\Python312\python.exe'
  cliPath: 'C:\Python312\Scripts\dxpdf_cli.py'
```

The `runner` field on a successful conversion names which candidate was used.

### `ModuleNotFoundError: No module named 'dxpdf'`

The engine is installed, but not into the interpreter the plugin invoked.
`python -m pip install dxpdf` installs into whatever `python` resolves to, and
the plugin may resolve a different interpreter (`pythonPath`, `DXPDF_PYTHON`,
`python3.exe`, a virtual environment). Print the interpreter that will be used
— `python -c "import sys; print(sys.executable)"` — and install into that one.

### Pagination is silently skipped

A conversion that reports
`paginationSkipped: "the pagination check failed: …"` is the usual symptom of a
missing `lxml`: Word-parity pagination runs `py/paginate.py`, which imports it.
Install `lxml` into the same interpreter as the engine
(`python -m pip install lxml`). Plain conversions do not need it.

### WPS documents fail

WPS Office writes schema-invalid constructs that the plugin repairs
automatically; see [WPS Office compatibility](#wps-office-compatibility). If
the engine rejects the document for a reason that is not a schema rejection, or
a named duplicate cannot be found, the tool error carries the engine's own
message.

### `python` vs `py -3` on Windows

The plugin searches `PATH` for `python.exe`, `python3.exe` and `python`; it does
not use the `py` launcher. If only the launcher is available, set `pythonPath`
(or `DXPDF_PYTHON`) to the interpreter it names —
`py -3 -c "import sys; print(sys.executable)"`. Conversion can still run through
an installed `dxpdf` shim without any interpreter on `PATH`, but the schema
repair and pagination passes are Python helpers, so they are skipped in that
case.

## Development

The profile can depend on this package through pnpm's `link:` protocol during
local work, so the source directory is authoritative: edit it and the installed
plugin is already current, with no reinstall step.

```json
"dsh-plugin-dxpdf": "link:../dsh-plugin-dxpdf"
```

This is deliberate. pnpm's `file:` protocol installs a *filtered copy* driven by
the `files` field, and `pnpm install` reports "Already up to date" without
refreshing that copy when only the package's contents changed — a silent drift
between source and install. `link:` removes the failure mode.

### Activating a change

The Harness composes a profile's plugin rows at startup. HMR reloads when the
ordered `dsh.profile.bundles` list changes, but a bundle whose package is not
yet on disk at that moment fails to import. After installing or editing a
bundle, **restart DeepSeek Harness**, or toggle the bundle on the Web sidebar's
**Plugins** page, which rewrites the bundles list and forces a recomposition.

## License

MIT — see [LICENSE](LICENSE).
