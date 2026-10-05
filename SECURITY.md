# Security policy

This plugin converts a Word document into a PDF. It is a local tool: it reads one
file, writes one file, and uses a scratch directory in between. The sections
below state exactly what it is allowed to touch, so a reader can check the claim
against the code instead of taking it on trust.

## Reporting a vulnerability

Open a private report through this repository's **Security → Report a
vulnerability** tab. Please include the version (`package.json`), the document or
command that triggers it, and the smallest reproduction you can manage.

## What the plugin can do

| Surface | What it does |
|---|---|
| Processes | Spawns the dxpdf engine and its own Python helpers with an **explicit argv array**. No shell is involved, so no value is ever parsed as a command. |
| Files read | The `.docx` the caller named, plus the engine and interpreter it was configured with. |
| Files written | The `.pdf` the caller named, and one `mkdtemp` directory under the system temp directory. Nothing else, at no point. |
| Network | None. No HTTP client, no socket, no DNS lookup, in JavaScript or in Python. |
| Credentials | None. No file named `.credentials.yaml` or `settings.yaml` is read, and no `*_API_KEY` / `*_TOKEN` variable is consulted. |
| Dynamic code | None. No `eval`, no `new Function`, no `vm`, no `os.system`. |
| Install time | Nothing runs. The package declares no `preinstall` / `install` / `postinstall` hook, and no `prepare` either. |
| Dependencies | Zero runtime dependencies. The profile gains this package and nothing else. |

## The environment a child receives

An allow-list decides it (`pythonChildEnv` in `lib/convert.js`): `PATH`,
`SystemRoot`, `TEMP`, the shell and program-directory variables, the locale
variables, plus anything named `PYTHON*` (the interpreter reads those) and
`DXPDF_*` (the plugin's own configuration).

The harness's environment also holds model API keys and session secrets. None of
them match the list, so none of them reach the converter — by construction
rather than by remembering to exclude them.

## What a document can and cannot do

The input is untrusted data. A `.docx` is a ZIP of XML, so the plugin treats it
as bytes and never executes anything from it:

* the engine parses XML with entity expansion left off, and the helpers rewrite
  XML text rather than importing or evaluating it;
* `py/normalize.py` and `py/omml.py` change only the elements they name and copy
  every other ZIP entry through byte for byte;
* a document cannot name a path: the input and output paths come from the
  caller, and the scratch directory is created by `mkdtemp` with a random name.

The realistic failure mode of a hostile document is therefore a crash or a slow
conversion, not code execution. A malformed archive is reported as a
non-zero exit with the engine's own message.

## Verifying this yourself

The claims above are enforced by `test/security.test.mjs`, which reads this
package's shipped files and fails if any of them stops holding:

```powershell
npm run test:security
```

For a third-party check of any installed plugin, `_tools_final/plugin_audit.py`
in the deployment applies the same checklist statically and reports the file and
line of anything it objects to.

## Scope and limits

* The PDF the plugin produces is written where the caller asked, so an output
  path inside a sensitive directory is the caller's decision, not a bypass.
* `py/paginate.py` is the only helper that imports a third-party module
  (`lxml`), and it is optional: a plain conversion does not need it.
* The engine itself (`dxpdf`, a Rust extension from PyPI) is out of scope here;
  this policy covers the plugin's own code.
