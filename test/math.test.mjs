/**
 * Math checks: an equation the engine cannot lay out must reach the page.
 *
 * dxpdf implements three OMML constructs and drops the whole subtree of every
 * other one *without failing*, so this defect cannot be caught by looking for a
 * conversion error — the PDF is produced and the formula is simply not in it.
 * `test/fixtures/math.docx` carries one equation per dropped construct, each
 * written with digit markers so the text layer can be asserted exactly.
 *
 * The suite therefore tests both directions: a plain conversion must be shown
 * to lose the markers (that is the bug), and a conversion through `convertDocx`
 * must be shown to keep them (that is the fix).
 */

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import {
	BUNDLED_CLI_PATH,
	convertDocx,
	lowerMathDocx,
	resolvePython,
} from "../lib/convert.js";

const PYTHON =
	(await resolvePython({
		env: process.env,
		pythonPath: process.env.DXPDF_TEST_PYTHON ?? process.env.DXPDF_PYTHON,
	})) ?? (process.platform === "win32" ? "python.exe" : "python3");
const CLI = process.env.DXPDF_TEST_CLI ?? process.env.DXPDF_CLI ?? BUNDLED_CLI_PATH;
const PYTHON_ENV = { DXPDF_PYTHON: PYTHON, DXPDF_CLI: CLI };
const fixture = (name) => fileURLToPath(new URL(`./fixtures/${name}`, import.meta.url));

/** Run one named check, reporting pass/fail without aborting the suite. */
async function check(label, body) {
	try {
		const detail = await body();
		console.log(`  PASS  ${label}${detail ? ` — ${detail}` : ""}`);
		return true;
	} catch (error) {
		console.log(`  FAIL  ${label} — ${error.message}`);
		return false;
	}
}

const digest = async (path) =>
	createHash("sha256").update(await readFile(path)).digest("hex");

/**
 * The text of a PDF, with the mathematical-italic code points the engine
 * substitutes for letters folded back to ASCII. Digits are never substituted,
 * which is why the fixture marks every equation with them.
 */
async function pdfText(path) {
	const { execFile } = await import("node:child_process");
	const script =
		"import sys,pymupdf;" +
		"d=pymupdf.open(sys.argv[1]);" +
		"sys.stdout.reconfigure(encoding='utf-8');" +
		"print(''.join(p.get_text() for p in d))";
	const text = await new Promise((resolve, reject) => {
		execFile(PYTHON, ["-c", script, path], { encoding: "utf-8" }, (error, stdout) =>
			error ? reject(error) : resolve(stdout),
		);
	});
	return text.replace(/[\u{1D400}-\u{1D7FF}]/gu, (char) => {
		const code = char.codePointAt(0);
		for (const [start, base] of [
			[0x1d400, 0x41], [0x1d41a, 0x61], [0x1d434, 0x41], [0x1d44e, 0x61],
			[0x1d468, 0x41], [0x1d482, 0x61], [0x1d4d0, 0x41], [0x1d4ea, 0x61],
		]) {
			if (code >= start && code < start + 26) return String.fromCharCode(base + (code - start));
		}
		return char;
	});
}

const results = [];
const scratch = await mkdtemp(join(tmpdir(), "dxpdf-math-test-"));
const maths = fixture("math.docx");
const plain = fixture("single-page.docx");

console.log("dxpdf math lowering");

results.push(
	await check("the engine really does drop these constructs", async () => {
		const out = join(scratch, "bug.pdf");
		await convertDocx({
			input: maths,
			output: out,
			env: PYTHON_ENV,
			lowerMath: false,
			matchWordPagination: false,
		});
		const text = await pdfText(out);
		for (const marker of ["22", "33", "111", "444", "55", "66", "77", "888"]) {
			assert.ok(!text.includes(marker), `marker ${marker} survived the plain engine`);
		}
		for (const symbol of ["\u221a", "\u2211", "["]) {
			assert.ok(!text.includes(symbol), `symbol ${symbol} reached the page unaided`);
		}
		assert.ok(text.includes("15"), "the construct the engine supports must render");
		return "markers 22/33/111/444/55/66/77/888 and \u221a \u2211 [ are absent without the pass";
	}),
);

results.push(
	await check("the same document converts with the equations intact", async () => {
		const out = join(scratch, "fixed.pdf");
		const facts = await convertDocx({ input: maths, output: out, env: PYTHON_ENV });
		assert.equal(facts.mathEquations, 15, "every equation should be seen");
		assert.equal(facts.mathLowered, 14, "every dropped one should be rewritten");
		const text = await pdfText(out);
		for (const marker of ["22", "33", "111", "444", "55", "66", "77", "888",
			"99", "10", "11", "12", "13", "14", "20", "21", "23", "24", "26"]) {
			assert.ok(text.includes(marker), `marker ${marker} is missing from the page`);
		}
		for (const symbol of ["\u221a", "\u2211", "[", "]", "\u02c6", "\u203e",
			"\u23df", "\u2082"]) {
			assert.ok(text.includes(symbol), `symbol ${symbol} is missing from the page`);
		}
		assert.ok(text.includes("sin"), "a function name should survive");
		return `${facts.mathLowered} of ${facts.mathEquations} equations lowered`;
	}),
);

results.push(
	await check("a document with nothing to lower is not rewritten", async () => {
		const out = join(scratch, "untouched.docx");
		const report = await lowerMathDocx({
			input: plain,
			output: out,
			python: PYTHON,
			timeoutMs: 60_000,
		});
		assert.equal(report.needed, false, "a clean document must need nothing");
		assert.equal(report.changed, false);
		assert.deepEqual(report.constructs, {});
		await assert.rejects(stat(out), "no copy should be written");
		return "no constructs, no file";
	}),
);

results.push(
	await check("the source document is never modified", async () => {
		const before = await digest(maths);
		await lowerMathDocx({
			input: maths,
			output: join(scratch, "again.docx"),
			python: PYTHON,
			timeoutMs: 60_000,
		});
		assert.equal(await digest(maths), before, "the fixture changed on disk");
		return "sha256 unchanged";
	}),
);

results.push(
	await check("--fraction linear trades the stack for a real subscript", async () => {
		const stacked = join(scratch, "stacked.docx");
		const linear = join(scratch, "linear.docx");
		await lowerMathDocx({
			input: maths, output: stacked, python: PYTHON, timeoutMs: 60_000,
			fraction: "native",
		});
		await lowerMathDocx({
			input: maths, output: linear, python: PYTHON, timeoutMs: 60_000,
			fraction: "linear",
		});
		const native = await pdfText(await convert(join(scratch, "stacked.pdf"), stacked));
		const flat = await pdfText(await convert(join(scratch, "linear.pdf"), linear));
		assert.ok(native.includes("\u2082"), "native keeps the script in Unicode form");
		assert.ok(!flat.includes("\u2082"), "linear must not need Unicode scripts");
		assert.ok(flat.includes("25"), "linear writes the script as its own run");
		return "native uses U+2082, linear writes 25 as a run";
	}),
);

results.push(
	await check("the options reach the helper and are reported back", async () => {
		const out = join(scratch, "options.docx");
		const report = await lowerMathDocx({
			input: maths,
			output: out,
			python: PYTHON,
			timeoutMs: 60_000,
			fraction: "linear",
			radical: "parens",
			scripts: "brackets",
		});
		assert.equal(report.fraction, "linear");
		assert.equal(report.radical, "parens");
		assert.equal(report.scripts, "brackets");
		assert.ok(report.lowered > 0);
		assert.deepEqual(report.warnings, []);
		return `fraction=${report.fraction} radical=${report.radical} scripts=${report.scripts}`;
	}),
);

/** Convert a document that the math pass already prepared. */
async function convert(output, input) {
	await convertDocx({
		input,
		output,
		env: PYTHON_ENV,
		lowerMath: false,
		matchWordPagination: false,
	});
	return output;
}

await rm(scratch, { recursive: true, force: true });

const passed = results.filter(Boolean).length;
console.log(`\n${passed}/${results.length} checks passed`);
process.exitCode = passed === results.length ? 0 : 1;
