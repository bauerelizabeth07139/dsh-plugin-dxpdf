/**
 * Headless checks for the dxpdf plugin core. These exercise the real engine on
 * real files — no Harness, no mocks — so a passing run means the tool body
 * would work inside the host process too.
 *
 * Usage: node test/convert.test.mjs [single-page.docx] [multi-page.docx]
 *
 * Both arguments default to the fixtures committed beside this file.
 */

import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import {
	BUNDLED_CLI_PATH,
	convertDocx,
	countPdfPages,
	resolvePython,
	resolveRunner,
} from "../lib/convert.js";

// Machine-local engine locations, so the suite runs anywhere. Each one falls
// back to the machine variables the plugin itself reads, then to the shim this
// package ships — the one that only needs `python -m pip install dxpdf`.
const PYTHON =
	(await resolvePython({
		env: process.env,
		pythonPath: process.env.DXPDF_TEST_PYTHON ?? process.env.DXPDF_PYTHON,
	})) ?? (process.platform === "win32" ? "python.exe" : "python3");
const CLI = process.env.DXPDF_TEST_CLI ?? process.env.DXPDF_CLI ?? BUNDLED_CLI_PATH;
const SHIM = process.env.DXPDF_TEST_SHIM ?? process.env.DXPDF_EXE ?? "";

const fixture = (name) =>
	fileURLToPath(new URL(`./fixtures/${name}`, import.meta.url));

const [, , singleArg, multiArg] = process.argv;
const singlePage = singleArg ?? fixture("single-page.docx");
const multiPage = multiArg ?? fixture("multi-page.docx");

/** Run one named check, reporting pass/fail without aborting the suite. */
async function check(label, body) {
	try {
		const detail = await body();
		if (typeof detail === "string" && detail.startsWith("SKIP")) {
			skipped++;
			console.log(`  SKIP  ${label}${detail.length > 4 ? ` — ${detail.slice(5)}` : ""}`);
			return true;
		}
		console.log(`  PASS  ${label}${detail ? ` — ${detail}` : ""}`);
		return true;
	} catch (error) {
		console.log(`  FAIL  ${label} — ${error.message}`);
		return false;
	}
}

const results = [];
let skipped = 0;
const scratch = await mkdtemp(join(tmpdir(), "dxpdf-plugin-test-"));

console.log("dxpdf plugin core");

results.push(
	await check("resolves the runner from the declared DXPDF_* variables", async () => {
		const runner = await resolveRunner({
			env: { DXPDF_PYTHON: PYTHON, DXPDF_CLI: CLI },
		});
		assert.equal(runner.kind, "python");
		assert.equal(runner.command, PYTHON);
		return `${runner.kind} via ${runner.label}`;
	}),
);

results.push(
	await check("resolves a runner from PATH alone (no DXPDF_* variables)", async () => {
		const runner = await resolveRunner({ env: { PATH: process.env.PATH } });
		assert.ok(runner.command, "expected a runner");
		return `${runner.kind} via ${runner.label}`;
	}),
);

results.push(
	await check("resolves and uses the .cmd shim strategy", async () => {
		if (SHIM === "") {
			return "SKIP this machine has no dxpdf shim; set DXPDF_TEST_SHIM to one";
		}
		const runner = await resolveRunner({ env: {}, executable: SHIM });
		assert.equal(runner.kind, "shim");
		const out = join(scratch, "via-shim.pdf");
		const facts = await convertDocx({
			input: singlePage,
			output: out,
			env: {},
			executable: SHIM,
		});
		assert.ok(facts.bytes > 1000, `expected a real PDF, got ${facts.bytes} bytes`);
		return `${facts.bytes} bytes through cmd.exe`;
	}),
);

results.push(
	await check("converts a one-page DOCX and reads the page count back", async () => {
		const out = join(scratch, "single.pdf");
		const facts = await convertDocx({
			input: singlePage,
			output: out,
			env: { DXPDF_PYTHON: PYTHON, DXPDF_CLI: CLI },
		});
		const head = (await readFile(out)).subarray(0, 5).toString("latin1");
		assert.equal(head, "%PDF-");
		assert.equal(facts.pages, 1);
		return `${facts.bytes} bytes, ${facts.pages} page, ${facts.elapsedMs} ms`;
	}),
);

results.push(
	await check("converts a multi-page DOCX and agrees with an independent count", async () => {
		const out = join(scratch, "multi.pdf");
		const facts = await convertDocx({
			input: multiPage,
			output: out,
			env: { DXPDF_PYTHON: PYTHON, DXPDF_CLI: CLI },
		});
		const bytes = await readFile(out);
		const independent = (bytes.toString("latin1").match(/\/Type\s*\/Page(?![s])/g) ?? []).length;
		assert.equal(facts.pages, independent);
		assert.ok(facts.pages > 1, `expected more than one page, got ${facts.pages}`);
		return `${facts.pages} pages, ${facts.bytes} bytes`;
	}),
);

results.push(
	await check("honours image_dpi", async () => {
		const out = join(scratch, "dpi.pdf");
		const facts = await convertDocx({
			input: singlePage,
			output: out,
			imageDpi: 96,
			env: { DXPDF_PYTHON: PYTHON, DXPDF_CLI: CLI },
		});
		assert.ok(facts.bytes > 0);
		return `image_dpi=96 produced ${facts.bytes} bytes`;
	}),
);

results.push(
	await check("rejects an out-of-range image_dpi before spawning", async () => {
		await assert.rejects(
			() =>
				convertDocx({
					input: singlePage,
					output: join(scratch, "bad.pdf"),
					imageDpi: 99_999,
					env: { DXPDF_PYTHON: PYTHON, DXPDF_CLI: CLI },
				}),
			/invalid image_dpi/,
		);
		return "threw as expected";
	}),
);

results.push(
	await check("reports a missing input as a conversion failure", async () => {
		await assert.rejects(
			() =>
				convertDocx({
					input: join(scratch, "does-not-exist.docx"),
					output: join(scratch, "nope.pdf"),
					env: { DXPDF_PYTHON: PYTHON, DXPDF_CLI: CLI },
				}),
			/exited with code 1|no such file/i,
		);
		return "threw as expected";
	}),
);

results.push(
	await check("countPdfPages returns undefined for non-PDF bytes", async () => {
		assert.equal(countPdfPages(Buffer.from("not a pdf at all")), undefined);
		return "undefined";
	}),
);

results.push(
	await check("overwrites an existing output file", async () => {
		const out = join(scratch, "overwrite.pdf");
		const env = { DXPDF_PYTHON: PYTHON, DXPDF_CLI: CLI };
		const first = await convertDocx({ input: singlePage, output: out, env });
		const second = await convertDocx({ input: multiPage, output: out, env });
		assert.ok(second.bytes !== first.bytes || second.pages !== first.pages);
		assert.equal((await stat(out)).isFile(), true);
		return `${first.pages} page -> ${second.pages} pages in place`;
	}),
);

results.push(
	await check("falls back to the CLI shim this package ships", async () => {
		const runner = await resolveRunner({ env: { PATH: "" }, pythonPath: PYTHON });
		assert.equal(runner.kind, "python");
		assert.equal(runner.prefixArgs[0], BUNDLED_CLI_PATH);
		return `${runner.label} + ${PYTHON}`;
	}),
);

results.push(
	await check("converts through the shipped shim when nothing is on PATH", async () => {
		const out = join(scratch, "bundled.pdf");
		const facts = await convertDocx({
			input: singlePage,
			output: out,
			env: { PATH: "" },
			pythonPath: PYTHON,
		});
		assert.ok(facts.bytes > 1000, `expected a real PDF, got ${facts.bytes} bytes`);
		assert.equal(facts.runner, "bundled py/dxpdf_cli.py");
		return `${facts.bytes} bytes via ${facts.runner}`;
	}),
);

await rm(scratch, { recursive: true, force: true });

const passed = results.filter(Boolean).length;
console.log(
	`\n${passed}/${results.length} checks passed${skipped ? `, ${skipped} skipped` : ""}`,
);
process.exit(passed === results.length ? 0 : 1);
