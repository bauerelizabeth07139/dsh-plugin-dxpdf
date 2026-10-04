/**
 * Repair checks: the plugin must convert a WPS-authored document that dxpdf
 * rejects outright, without the caller preparing anything.
 *
 * `test/fixtures/wps-defects.docx` carries one of each defect (duplicate
 * `w:sz`, duplicate `w:pPr`, an equation wrapped in a run). The base fixture
 * is well-formed, so it exercises the "no repair needed" path.
 */

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import {
	BUNDLED_CLI_PATH,
	convertDocx,
	duplicateHints,
	normalizeDocx,
	resolvePython,
} from "../lib/convert.js";

// Machine-local engine locations, so the suite runs anywhere: the test
// overrides, then the machine variables the plugin reads, then the shim this
// package ships (which only needs `python -m pip install dxpdf`).
const PYTHON =
	process.env.DXPDF_TEST_PYTHON ??
	process.env.DXPDF_PYTHON ??
	(process.platform === "win32" ? "python.exe" : "python3");
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

const results = [];
const scratch = await mkdtemp(join(tmpdir(), "dxpdf-repair-test-"));
const broken = fixture("wps-defects.docx");
const healthy = fixture("single-page.docx");

console.log("dxpdf document repair");

results.push(
	await check("duplicateHints reads the field the engine rejected", async () => {
		const message =
			"dxpdf: conversion failed: Parse error: failed to deserialize XML: duplicate field `sz`";
		assert.deepEqual(duplicateHints(message), ["sz"]);
		assert.deepEqual(duplicateHints("nothing useful here"), []);
		assert.deepEqual(
			duplicateHints("duplicate field `sz` ... duplicate field `pPr` ... duplicate field `sz`"),
			["sz", "pPr"],
		);
		return "sz, pPr extracted without repeats";
	}),
);

results.push(
	await check("a well-formed document converts with no repair pass", async () => {
		const out = join(scratch, "clean.pdf");
		const facts = await convertDocx({ input: healthy, output: out, env: PYTHON_ENV });
		assert.equal(facts.normalized, false);
		assert.equal(facts.edits, 0);
		assert.equal(facts.rounds, 0);
		assert.equal(facts.pages, 1);
		return `${facts.bytes} bytes, no repair`;
	}),
);

results.push(
	await check("the raw engine rejects the WPS fixture on its own", async () => {
		// Proves the fixture is genuinely broken and that it is the repair — not
		// luck — that makes convertDocx succeed below.
		const proc = spawnSync(
			PYTHON,
			[
				CLI,
				broken,
				"-o",
				join(scratch, "raw.pdf"),
			],
			{ encoding: "utf8" },
		);
		assert.notEqual(proc.status, 0, "expected a non-zero exit from the raw engine");
		assert.match(proc.stderr, /duplicate field|deserialize/i);
		return (proc.stderr.trim().split("\n").pop() ?? "").slice(0, 96);
	}),
);

results.push(
	await check("convertDocx repairs and converts the WPS fixture", async () => {
		const out = join(scratch, "repaired.pdf");
		const facts = await convertDocx({ input: broken, output: out, env: PYTHON_ENV });
		assert.equal(facts.normalized, true, "expected the repair pass to run");
		// The engine names only `sz`, but the structural repairs — equation
		// unwrap, nested-run unwrap, property reorder — need no permission and
		// ride along in the same pass. The duplicate `w:pPr` is schema-invalid
		// yet tolerated by dxpdf, so a targeted repair leaves it alone.
		assert.ok(facts.edits >= 4, `expected at least 4 edits, got ${facts.edits}`);
		assert.ok(facts.rounds >= 1, "expected at least one repair round");
		const head = (await readFile(out)).subarray(0, 5).toString("latin1");
		assert.equal(head, "%PDF-");
		assert.equal(facts.pages, 1);
		return `${facts.edits} edits over ${facts.rounds} round(s), ${facts.bytes} bytes`;
	}),
);

results.push(
	await check("the repair pass leaves the source document untouched", async () => {
		const before = await digest(broken);
		await convertDocx({ input: broken, output: join(scratch, "again.pdf"), env: PYTHON_ENV });
		assert.equal(await digest(broken), before);
		return "source digest unchanged";
	}),
);

results.push(
	await check("normalizeDocx reports every defect class", async () => {
		const repaired = join(scratch, "report.docx");
		const report = await normalizeDocx({
			input: broken,
			output: repaired,
			python: PYTHON,
			timeoutMs: 60_000,
		});
		assert.equal(report.changed, true);
		// One `sz` deletion, one duplicate `pPr` deletion, two edits for the
		// misplaced `pPr` (insert at the front, delete where it stood), one
		// nested-run unwrap, one equation unwrap.
		assert.equal(report.edits, 6, `expected 6 edits, got ${report.edits}`);
		assert.deepEqual(report.parts, {
			"word/document.xml": 5,
			"word/styles.xml": 1,
		});
		const kinds = new Set(report.repairs.map((entry) => entry.kind));
		for (const kind of [
			"duplicate-property",
			"misplaced-property",
			"nested-run",
			"math-in-run",
		]) {
			assert.ok(kinds.has(kind), `no ${kind} repair recorded`);
		}
		assert.equal((await stat(repaired)).isFile(), true);
		return `${report.edits} edits: ${[...kinds].sort().join(", ")}`;
	}),
);

results.push(
	await check("a repaired copy is a valid DOCX the engine then accepts", async () => {
		const repaired = join(scratch, "valid.docx");
		await normalizeDocx({
			input: broken,
			output: repaired,
			python: PYTHON,
			timeoutMs: 60_000,
		});
		const facts = await convertDocx({
			input: repaired,
			output: join(scratch, "valid.pdf"),
			env: PYTHON_ENV,
		});
		assert.equal(facts.normalized, false, "a repaired copy must not need repairing again");
		assert.equal(facts.pages, 1);
		return "converted with no further repair";
	}),
);

results.push(
	await check("both duplicate policies produce a convertible document", async () => {
		const sizes = [];
		for (const policy of ["first", "last"]) {
			const repaired = join(scratch, `${policy}.docx`);
			const report = await normalizeDocx({
				input: broken,
				output: repaired,
				python: PYTHON,
				timeoutMs: 60_000,
				policy,
			});
			assert.equal(report.policy, policy);
			assert.ok(report.edits >= 6, `expected at least 6 edits, got ${report.edits}`);
			const facts = await convertDocx({
				input: repaired,
				output: join(scratch, `${policy}.pdf`),
				env: PYTHON_ENV,
			});
			sizes.push(`${policy}=${facts.bytes}B`);
		}
		return sizes.join(", ");
	}),
);

results.push(
	await check("resolvePython finds the interpreter behind DXPDF_PYTHON", async () => {
		const python = await resolvePython({ env: PYTHON_ENV });
		assert.equal(python, PYTHON);
		return python;
	}),
);

await rm(scratch, { recursive: true, force: true });

const passed = results.filter(Boolean).length;
console.log(`\n${passed}/${results.length} checks passed`);
process.exit(passed === results.length ? 0 : 1);
