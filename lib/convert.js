/**
 * Conversion core for the `dxpdf_convert` tool — everything that can be
 * reasoned about and tested without a running Harness.
 *
 * The `dxpdf` wheel on PyPI ships only the compiled extension module plus a
 * `convert_file` function; it has no console script. This module therefore
 * resolves a *runner* — an interpreter plus the CLI shim, or the shim itself —
 * from the machine variables the install declared, then drives it with an
 * explicit argv (never a shell string the model could influence).
 *
 * dxpdf parses OOXML with strict schemas, so a document WPS Office wrote can
 * be rejected even though Word opens it happily. When that happens the engine
 * is handed a repaired copy instead — see `py/normalize.py` — so the caller
 * never has to know the difference.
 *
 * Runner resolution order, first candidate that exists on disk wins:
 *   1. explicit `pythonPath` + `cliPath` (plugin configuration)
 *   2. `%DXPDF_PYTHON%` + `%DXPDF_CLI%` (the declared machine variables)
 *   3. `%DXPDF_EXE%` (the declared `dxpdf.cmd` shim, run through `cmd.exe`)
 *   4. `dxpdf` found on `PATH` (run through `cmd.exe`)
 *   5. a Python on `PATH` paired with the CLI shim this package ships
 *      (`py/dxpdf_cli.py`), so `pip install dxpdf` is the only prerequisite
 *   6. a Python on `PATH` paired with `dxpdf_cli.py` on `PATH`
 *
 * @module dsh-plugin-dxpdf/convert
 */

import { spawn } from "node:child_process";
import { constants } from "node:fs";
import { access, mkdir, mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { delimiter, dirname, isAbsolute, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

/** Embedded-image resolution dxpdf defaults to, matching Word. */
export const DEFAULT_IMAGE_DPI = 220;
/** The range dxpdf accepts for `--image-dpi`. */
export const IMAGE_DPI_RANGE = Object.freeze([1, 2400]);
/** Default wall-clock budget for one conversion. */
export const DEFAULT_TIMEOUT_MS = 120_000;

/** The repair helper shipped beside this module. */
export const NORMALIZER_PATH = fileURLToPath(
	new URL("../py/normalize.py", import.meta.url),
);

/** The pagination helper shipped beside this module. */
export const PAGINATOR_PATH = fileURLToPath(
	new URL("../py/paginate.py", import.meta.url),
);

/**
 * The CLI shim shipped beside this module. The `dxpdf` wheel declares no
 * console script, so a machine that only ran `pip install dxpdf` has no
 * `dxpdf` command; this copy of the shim is what makes that machine work.
 */
export const BUNDLED_CLI_PATH = fileURLToPath(
	new URL("../py/dxpdf_cli.py", import.meta.url),
);

/**
 * Default budget for the pagination pass, which runs the engine once per
 * section. It has to comfortably exceed one engine run per parity break, or a
 * document that is inside `maxPaginateSections` fails on the clock instead of
 * converging: at roughly 20 s a probe, 48 breaks need about 16 minutes.
 */
export const DEFAULT_PAGINATE_TIMEOUT_MS = 2_700_000;

/** Refuse the pagination pass beyond this many parity breaks, to bound the cost. */
export const DEFAULT_MAX_PAGINATE_SECTIONS = 48;

/** Windows `cmd.exe`, the only way to invoke a `.cmd`/`.bat` shim. */
const COMSPEC = process.env.ComSpec ?? process.env.COMSPEC ?? "cmd.exe";

/**
 * Engine messages that mean "this document's XML is not schema-valid", as
 * opposed to a missing file, a timeout, or an I/O fault. Only these justify
 * paying for a repair pass and a second conversion.
 */
const SCHEMA_REJECTION = /failed to deserialize|duplicate field|unknown variant|parse error/i;

/**
 * The property element names the engine has already rejected as duplicates.
 * `failed to deserialize XML: duplicate field \`sz\`` names exactly one, which
 * is far better evidence than any built-in list of "should appear once"
 * elements: dxpdf's own schema is what decides.
 * @param {string} message - the engine's combined output.
 * @returns {string[]} the names it called duplicates, in order.
 */
export function duplicateHints(message) {
	const names = [];
	for (const match of message.matchAll(/duplicate field\s+`([^`]+)`/g)) {
		if (!names.includes(match[1])) names.push(match[1]);
	}
	return names;
}

/**
 * How many repair rounds to attempt. Each round is cheap: the engine rejects a
 * bad document while parsing, before doing any layout or painting.
 */
const MAX_REPAIR_ROUNDS = 6;

/**
 * Whether a path exists as a regular file.
 * @param {string | undefined | null} path - candidate path.
 * @returns {Promise<boolean>} true when the path is an accessible file.
 */
async function isFile(path) {
	if (typeof path !== "string" || path.length === 0) return false;
	try {
		await access(path, constants.F_OK);
		return (await stat(path)).isFile();
	} catch {
		return false;
	}
}

/**
 * Quote one token for a `cmd.exe` command line. Windows filenames cannot
 * contain a double quote, so doubling is belt-and-braces rather than load-
 * bearing; it keeps the helper correct for any caller that passes one.
 * @param {string} token - raw argument.
 * @returns {string} the quoted token.
 */
function quoteForCmd(token) {
	return `"${String(token).replaceAll('"', '""')}"`;
}

/**
 * Search `PATH` for the first of the given file names that exists.
 * @param {string[]} names - candidate file names, in preference order.
 * @param {Record<string, string | undefined>} env - environment to read.
 * @returns {Promise<string | undefined>} the first existing absolute path.
 */
async function findOnPath(names, env) {
	const raw = env.PATH ?? env.Path ?? env.path ?? "";
	for (const dir of raw.split(delimiter)) {
		if (dir.length === 0) continue;
		for (const name of names) {
			const candidate = join(dir, name);
			if (await isFile(candidate)) return candidate;
		}
	}
	return undefined;
}

/**
 * A runner is `spawn`-ready: a concrete executable plus the argv prefix that
 * selects the CLI, and the strategy that knows how to add per-call arguments.
 * @typedef {object} Runner
 * @property {"python" | "shim"} kind - which strategy produced this runner.
 * @property {string} command - executable to spawn.
 * @property {string[]} prefixArgs - argv before the per-call arguments.
 * @property {string} [shim] - the shim path, for the `shim` kind.
 * @property {string} [python] - the interpreter, when one is known.
 * @property {string} label - human-readable provenance for error messages.
 */

/**
 * Resolve a usable dxpdf runner from configuration and the declared machine
 * variables.
 *
 * The machine variables are the fast path, but they are only inherited by
 * processes started after they were declared — a harness that was already
 * running when they were set will not see them. The PATH search and the
 * conventional-interpreter fallbacks below therefore make the plugin work
 * without them, so a restart is not required for correctness.
 *
 * @param {object} [options] - resolution inputs.
 * @param {Record<string, string | undefined>} [options.env] - environment to read.
 * @param {string} [options.pythonPath] - explicit interpreter override.
 * @param {string} [options.cliPath] - explicit CLI-script override.
 * @param {string} [options.executable] - explicit shim/executable override.
 * @returns {Promise<Runner>} the first runner that exists.
 * @throws {Error} when no candidate resolves.
 */
export async function resolveRunner(options = {}) {
	const env = options.env ?? process.env;
	const tried = [];

	const considerPython = async (pythonPath, cliPath, label) => {
		if (!pythonPath || !cliPath) return undefined;
		tried.push(`${label}: ${pythonPath} + ${cliPath}`);
		if ((await isFile(pythonPath)) && (await isFile(cliPath))) {
			return {
				kind: "python",
				command: pythonPath,
				prefixArgs: [cliPath],
				python: pythonPath,
				label,
			};
		}
		return undefined;
	};

	const explicit = await considerPython(
		options.pythonPath,
		options.cliPath,
		"configured pythonPath/cliPath",
	);
	if (explicit) return explicit;

	const declared = await considerPython(
		env.DXPDF_PYTHON,
		env.DXPDF_CLI,
		"DXPDF_PYTHON/DXPDF_CLI machine variables",
	);
	if (declared) return declared;

	// The shim, explicit first and then discovered on PATH so a machine whose
	// environment was never refreshed still works.
	const shimCandidates = [
		[options.executable, "configured executable"],
		[env.DXPDF_EXE, "DXPDF_EXE machine variable"],
		[await findOnPath(["dxpdf.cmd", "dxpdf.bat", "dxpdf.exe"], env), "dxpdf on PATH"],
	];
	for (const [shim, label] of shimCandidates) {
		if (!shim) continue;
		tried.push(`${label}: ${shim}`);
		if (await isFile(shim)) {
			const python = await resolvePython({ env, pythonPath: options.pythonPath });
			return { kind: "shim", command: COMSPEC, prefixArgs: [], shim, python, label };
		}
	}

	// The shim this package ships: any Python plus our own copy of the CLI, so a
	// machine whose only preparation was `python -m pip install dxpdf` converts
	// without a separately installed `dxpdf` command.
	const bundledPython = await resolvePython({ env, pythonPath: options.pythonPath });
	if (bundledPython && (await isFile(BUNDLED_CLI_PATH))) {
		tried.push(`bundled shim: ${bundledPython} + ${BUNDLED_CLI_PATH}`);
		return {
			kind: "python",
			command: bundledPython,
			prefixArgs: [BUNDLED_CLI_PATH],
			python: bundledPython,
			label: "bundled py/dxpdf_cli.py",
		};
	}

	// Last resort: any Python on PATH paired with a CLI script on PATH.
	const pythonOnPath = await findOnPath(["python.exe", "python3.exe", "python"], env);
	const cliOnPath = await findOnPath(["dxpdf_cli.py"], env);
	const discovered = await considerPython(pythonOnPath, cliOnPath, "discovered on PATH");
	if (discovered) return discovered;

	throw new Error(
		"dxpdf is not installed where this plugin can find it. Tried " +
			`${tried.join("; ") || "nothing"}. Install it with ` +
			'"python -m pip install dxpdf", or point the plugin at it with the ' +
			"pythonPath/cliPath options.",
	);
}

/**
 * Resolve a Python interpreter, used to run the repair helper.
 * @param {object} [options] - resolution inputs.
 * @param {Record<string, string | undefined>} [options.env] - environment to read.
 * @param {string} [options.pythonPath] - explicit interpreter override.
 * @returns {Promise<string | undefined>} the interpreter path, when one exists.
 */
export async function resolvePython(options = {}) {
	const env = options.env ?? process.env;
	for (const candidate of [
		options.pythonPath,
		env.DXPDF_PYTHON,
		await findOnPath(["python.exe", "python3.exe", "python"], env),
	]) {
		if (candidate && (await isFile(candidate))) return candidate;
	}
	return undefined;
}

/**
 * Build the full argv for one conversion under a runner.
 * @param {Runner} runner - resolved runner.
 * @param {string[]} args - per-call CLI arguments.
 * @returns {string[]} the argv to spawn.
 */
function argvFor(runner, args) {
	if (runner.kind === "python") return [...runner.prefixArgs, ...args];
	// `cmd /d /s /c "<command line>"` — /d skips AutoRun, /s keeps the outer
	// quotes literal so the shim's own `%*` forwarding sees clean arguments.
	const line = [runner.shim, ...args].map(quoteForCmd).join(" ");
	return ["/d", "/s", "/c", `"${line}"`];
}

/**
 * Best-effort page count from PDF bytes. dxpdf emits uncompressed page
 * dictionaries, so counting `/Type /Page` (never `/Pages`) is reliable for its
 * own output; object streams in foreign PDFs would undercount, which is why
 * the caller treats an absent count as "unknown" rather than zero.
 * @param {Buffer} bytes - the produced PDF.
 * @returns {number | undefined} the page count when at least one page is found.
 */
export function countPdfPages(bytes) {
	const matches = bytes.toString("latin1").match(/\/Type\s*\/Page(?![s])/g);
	return matches === null ? undefined : matches.length;
}

/**
 * Run a child process to completion, capturing both streams.
 * @param {string} command - executable.
 * @param {string[]} args - argv.
 * @param {object} options - execution controls.
 * @param {AbortSignal} [options.signal] - caller cancellation.
 * @param {number} options.timeoutMs - wall-clock budget.
 * @param {string} [options.cwd] - working directory.
 * @param {boolean} [options.verbatim] - pass `windowsVerbatimArguments`, required
 *   when the last argument is a complete `cmd.exe` command line that must not
 *   be re-quoted by Node.
 * @param {Record<string, string | undefined>} [options.childEnv] - environment for the child.
 * @returns {Promise<{ exitCode: number, stdout: string, stderr: string, timedOut: boolean }>} outcome.
 */
function run(command, args, { signal, timeoutMs, cwd, verbatim = false, childEnv }) {
	return new Promise((resolvePromise, rejectPromise) => {
		let child;
		try {
			child = spawn(command, args, {
				cwd,
				windowsHide: true,
				stdio: ["ignore", "pipe", "pipe"],
				...(verbatim ? { windowsVerbatimArguments: true } : {}),
				...(childEnv === undefined ? {} : { env: childEnv }),
				...(signal === undefined ? {} : { signal }),
			});
		} catch (error) {
			rejectPromise(error);
			return;
		}

		let stdout = "";
		let stderr = "";
		let timedOut = false;
		const timer = setTimeout(() => {
			timedOut = true;
			child.kill();
		}, timeoutMs);
		timer.unref?.();

		child.stdout.setEncoding("utf8");
		child.stderr.setEncoding("utf8");
		child.stdout.on("data", (chunk) => {
			stdout += chunk;
		});
		child.stderr.on("data", (chunk) => {
			stderr += chunk;
		});
		child.on("error", (error) => {
			clearTimeout(timer);
			rejectPromise(error);
		});
		child.on("close", (code) => {
			clearTimeout(timer);
			resolvePromise({ exitCode: code ?? -1, stdout, stderr, timedOut });
		});
	});
}

/** The environment a piped Python child needs to emit UTF-8 rather than cp936. */
function pythonChildEnv() {
	return { ...process.env, PYTHONIOENCODING: "utf-8" };
}

/**
 * Describe a failed engine run for the model.
 * @param {{ exitCode: number, stdout: string, stderr: string, timedOut: boolean }} result - outcome.
 * @param {string} input - the document that was attempted.
 * @param {number} timeoutMs - the budget that applied.
 * @returns {string} the message.
 */
function describeFailure(result, input, timeoutMs) {
	if (result.timedOut) {
		return (
			`dxpdf timed out after ${timeoutMs} ms converting ${input}. ` +
			"Raise the plugin's timeoutMs, or convert a smaller document."
		);
	}
	const detail = (result.stderr.trim() || result.stdout.trim() || "no output")
		.split("\n")
		.slice(-6)
		.join("\n");
	return `dxpdf exited with code ${result.exitCode}:\n${detail}`;
}

/**
 * Resolve a model-supplied path against the session working directory.
 * @param {string} value - the path as given.
 * @param {string} [cwd] - session working directory.
 * @returns {string} an absolute path.
 */
export function resolvePath(value, cwd) {
	return isAbsolute(value) ? value : resolve(cwd ?? process.cwd(), value);
}

/**
 * Run the repair helper over one document.
 * @param {object} options - repair request.
 * @param {string} options.input - the rejected document.
 * @param {string} options.output - path for the repaired copy.
 * @param {string} options.python - interpreter to run the helper with.
 * @param {AbortSignal} [options.signal] - caller cancellation.
 * @param {number} options.timeoutMs - wall-clock budget.
 * @param {string} [options.policy] - which duplicate occurrence to keep.
 * @param {string[]} [options.only] - property names to deduplicate; omit to
 *   sweep the helper's built-in singleton list instead.
 * @returns {Promise<{ changed: boolean, mode: string, edits: number, parts: Record<string, number>, repairs: object[] }>} the report.
 * @throws {Error} when the helper cannot run or reports a failure.
 */
export async function normalizeDocx({ input, output, python, signal, timeoutMs, policy, only }) {
	const args = [
		NORMALIZER_PATH,
		input,
		"--output",
		output,
		...(policy === undefined ? [] : ["--policy", policy]),
		...(only === undefined || only.length === 0 ? [] : ["--only", only.join(",")]),
	];
	const result = await run(python, args, {
		signal,
		timeoutMs,
		childEnv: pythonChildEnv(),
	});
	if (result.timedOut) {
		throw new Error(`the dxpdf repair helper timed out after ${timeoutMs} ms`);
	}
	if (result.exitCode !== 0) {
		throw new Error(
			`the dxpdf repair helper failed (exit ${result.exitCode}): ${result.stderr.trim() || "no output"}`,
		);
	}
	return JSON.parse(result.stdout);
}

/**
 * Ask the pagination helper which section breaks promise a page parity.
 *
 * This is the cheap gate in front of an expensive pass: the helper answers from
 * a byte scan of the main part for documents with no odd/even break at all, so
 * the common case costs one short process rather than one engine run per
 * section.
 *
 * @param {object} options - check request.
 * @param {string} options.input - the document to inspect.
 * @param {string} options.python - interpreter to run the helper with.
 * @param {AbortSignal} [options.signal] - caller cancellation.
 * @param {number} options.timeoutMs - wall-clock budget.
 * @returns {Promise<{boundaries: {paragraph: number, type: string}[]}>} the parity breaks.
 * @throws {Error} when the helper cannot run.
 */
export async function paginateCheck({ input, python, signal, timeoutMs }) {
	const result = await run(python, [PAGINATOR_PATH, input, "--check"], {
		signal,
		timeoutMs,
		childEnv: pythonChildEnv(),
	});
	if (result.timedOut) {
		throw new Error(`the dxpdf pagination check timed out after ${timeoutMs} ms`);
	}
	if (result.exitCode !== 0) {
		throw new Error(
			`the dxpdf pagination check failed (exit ${result.exitCode}): ${result.stderr.trim() || "no output"}`,
		);
	}
	return JSON.parse(result.stdout);
}

/**
 * Converge a document onto Word's pagination.
 *
 * dxpdf treats an odd/even section break as a plain `nextPage`, so the blank
 * pages Word inserts to reach the promised parity never appear. The helper
 * measures where each section ends and adds one `pageBreakBefore` filler per
 * boundary that would otherwise land on the wrong parity.
 *
 * @param {object} options - convergence request.
 * @param {string} options.input - the document to correct.
 * @param {string} options.output - path for the corrected copy.
 * @param {string} options.python - interpreter to run the helper with.
 * @param {AbortSignal} [options.signal] - caller cancellation.
 * @param {number} options.timeoutMs - wall-clock budget for the whole pass.
 * @returns {Promise<{fillers: number, probes: number, pagesBefore: number, pagesAfter: number, boundaries: object[]}>} the report.
 * @throws {Error} when the helper cannot run or reports a failure.
 */
export async function paginateDocx({ input, output, python, signal, timeoutMs }) {
	const result = await run(
		python,
		[PAGINATOR_PATH, input, "--output", output],
		{ signal, timeoutMs, childEnv: pythonChildEnv() },
	);
	if (result.timedOut) {
		throw new Error(`the dxpdf pagination pass timed out after ${timeoutMs} ms`);
	}
	if (result.exitCode !== 0) {
		throw new Error(
			`the dxpdf pagination pass failed (exit ${result.exitCode}): ${result.stderr.trim() || "no output"}`,
		);
	}
	return JSON.parse(result.stdout);
}

/**
 * Convert one DOCX file to PDF, repairing the document first when dxpdf's
 * strict OOXML reader rejects it.
 * @param {object} options - conversion request.
 * @param {string} options.input - source `.docx` path (already absolute).
 * @param {string} options.output - destination `.pdf` path (already absolute).
 * @param {number} [options.imageDpi] - embedded image resolution.
 * @param {number} [options.timeoutMs] - per-attempt wall-clock budget.
 * @param {AbortSignal} [options.signal] - caller cancellation.
 * @param {Record<string, string | undefined>} [options.env] - environment for runner resolution.
 * @param {string} [options.cwd] - working directory for the child.
 * @param {string} [options.pythonPath] - explicit interpreter override.
 * @param {string} [options.cliPath] - explicit CLI-script override.
 * @param {string} [options.executable] - explicit shim override.
 * @param {string} [options.policy] - which duplicate property occurrence wins.
 * @param {boolean} [options.matchWordPagination] - converge odd/even section
 *   breaks onto Word's pagination. Defaults to `true`; documents with no such
 *   break are detected cheaply and cost nothing.
 * @param {number} [options.paginateTimeoutMs] - budget for the pagination pass,
 *   which runs the engine once per parity section.
 * @param {number} [options.maxPaginateSections] - refuse the pass beyond this
 *   many parity breaks rather than spending an unbounded number of engine runs.
 * @returns {Promise<{bytes: number, elapsedMs: number, pages: number | undefined, runner: string, normalized: boolean, edits: number, repairs: object[], paginated: boolean, fillers: number, boundaries: number, paginationSkipped?: string}>} conversion facts.
 * @throws {Error} on a non-zero exit, a timeout, or a missing output file.
 */
export async function convertDocx(options) {
	const {
		input,
		output,
		imageDpi = DEFAULT_IMAGE_DPI,
		timeoutMs = DEFAULT_TIMEOUT_MS,
		signal,
		env,
		cwd,
		pythonPath,
		cliPath,
		executable,
		policy,
		matchWordPagination = true,
		paginateTimeoutMs = DEFAULT_PAGINATE_TIMEOUT_MS,
		maxPaginateSections = DEFAULT_MAX_PAGINATE_SECTIONS,
	} = options;

	const [minDpi, maxDpi] = IMAGE_DPI_RANGE;
	if (!Number.isFinite(imageDpi) || imageDpi < minDpi || imageDpi > maxDpi) {
		throw new Error(
			`invalid image_dpi: expected a number between ${minDpi} and ${maxDpi}, got ${JSON.stringify(imageDpi)}`,
		);
	}

	const runner = await resolveRunner({ env, pythonPath, cliPath, executable });
	const spawnOptions = {
		signal,
		timeoutMs,
		cwd,
		verbatim: runner.kind === "shim",
		childEnv: runner.kind === "python" ? pythonChildEnv() : undefined,
	};
	const attempt = (document) =>
		run(
			runner.command,
			argvFor(runner, [
				document,
				"--output",
				output,
				"--image-dpi",
				String(imageDpi),
				"--quiet",
			]),
			spawnOptions,
		);

	if (output !== input) {
		await mkdir(dirname(output), { recursive: true });
	}

	const started = Date.now();
	let working = input;
	let result = await attempt(working);
	let normalized = false;
	let edits = 0;
	let repairs = [];
	let rounds = 0;
	let paginated = false;
	let fillers = 0;
	let boundaries = 0;
	let paginationSkipped;

	// Every derived document lives in one scratch directory for the whole call,
	// so the repair pass and the pagination pass can hand work to each other.
	let scratch;
	const ensureScratch = async () => {
		scratch ??= await mkdtemp(join(tmpdir(), "dxpdf-convert-"));
		return scratch;
	};

	try {
		const python = runner.python ?? (await resolvePython({ env, pythonPath }));

		// A schema rejection is the one failure a repair can plausibly fix.
		// Anything else — a missing file, a timeout, an I/O fault — is reported
		// as it is.
		//
		// The engine names the field it choked on, so each round teaches the
		// repair pass one more name and it only ever removes properties dxpdf
		// itself said were duplicates. A rejected document fails during parsing,
		// so the extra rounds cost milliseconds.
		const rejection =
			result.exitCode !== 0 && !result.timedOut
				? `${result.stderr}\n${result.stdout}`
				: "";
		if (SCHEMA_REJECTION.test(rejection) && python !== undefined) {
			const directory = await ensureScratch();
			const hints = [];
			let failure = result;
			for (let round = 0; round < MAX_REPAIR_ROUNDS; round += 1) {
				const fresh = duplicateHints(`${failure.stderr}\n${failure.stdout}`).filter(
					(name) => !hints.includes(name),
				);
				if (fresh.length === 0) break;
				hints.push(...fresh);

				const repaired = join(directory, `repaired-${round}.docx`);
				let report;
				try {
					report = await normalizeDocx({
						input: working,
						output: repaired,
						python,
						signal,
						timeoutMs,
						policy,
						only: hints,
					});
				} catch (error) {
					throw new Error(
						`${describeFailure(failure, input, timeoutMs)}\n` +
							`The document is not schema-valid XML, and repairing it failed: ${error.message}`,
					);
				}
				if (!report.changed) {
					throw new Error(
						`${describeFailure(failure, input, timeoutMs)}\n` +
							`The engine named ${hints.map((h) => `\`${h}\``).join(", ")} as duplicated, but no such duplicate was found.`,
					);
				}
				normalized = true;
				edits = report.edits;
				repairs = report.repairs ?? [];
				rounds = round + 1;

				const retry = await attempt(repaired);
				if (retry.exitCode === 0 && !retry.timedOut) {
					working = repaired;
					result = retry;
					break;
				}
				failure = retry;
				result = retry;
			}
			if (result.exitCode !== 0 || result.timedOut) {
				throw new Error(
					`${describeFailure(result, input, timeoutMs)}\n` +
						`Repaired ${edits} schema violation(s) over ${rounds} round(s) (${hints.join(", ")}), and dxpdf still rejected the document.`,
				);
			}
		}

		// Word and dxpdf disagree about odd/even section breaks, which dxpdf
		// treats as a plain nextPage. Only a document that actually carries one
		// pays for this: the check answers from a byte scan first.
		if (matchWordPagination && python !== undefined && result.exitCode === 0) {
			let check;
			try {
				check = await paginateCheck({ input: working, python, signal, timeoutMs });
			} catch (error) {
				paginationSkipped = `the pagination check failed: ${error.message}`;
			}
			if (check !== undefined && check.boundaries.length > 0) {
				boundaries = check.boundaries.length;
				if (boundaries > maxPaginateSections) {
					paginationSkipped =
						`${boundaries} parity section breaks exceed maxPaginateSections ` +
						`(${maxPaginateSections}); raise it to converge this document`;
				} else {
					const directory = await ensureScratch();
					const paged = join(directory, "paginated.docx");
					try {
						const report = await paginateDocx({
							input: working,
							output: paged,
							python,
							signal,
							timeoutMs: paginateTimeoutMs,
						});
						if (report.fillers > 0) {
							const retry = await attempt(paged);
							if (retry.exitCode === 0 && !retry.timedOut) {
								working = paged;
								result = retry;
								paginated = true;
								fillers = report.fillers;
							} else {
								paginationSkipped =
									"the engine rejected the pagination-corrected copy";
							}
						}
					} catch (error) {
						paginationSkipped = `the pagination pass failed: ${error.message}`;
					}
				}
			}
		}
	} finally {
		if (scratch !== undefined) await rm(scratch, { recursive: true, force: true });
	}

	const elapsedMs = Date.now() - started;
	if (result.timedOut || result.exitCode !== 0) {
		throw new Error(describeFailure(result, input, timeoutMs));
	}

	const produced = await stat(output).catch(() => undefined);
	if (produced === undefined || !produced.isFile()) {
		throw new Error(
			`dxpdf reported success but wrote no file at ${output}. stderr: ${result.stderr.trim() || "(empty)"}`,
		);
	}

	const pages = countPdfPages(await readFile(output));
	return {
		bytes: produced.size,
		elapsedMs,
		pages,
		runner: runner.label,
		normalized,
		edits,
		rounds,
		repairs,
		paginated,
		fillers,
		boundaries,
		...(paginationSkipped === undefined ? {} : { paginationSkipped }),
	};
}
