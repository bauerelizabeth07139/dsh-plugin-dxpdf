/**
 * Pagination checks: a document with an odd-page section break must come out of
 * the tool with the blank page Word inserts, and documents without one must not
 * pay for the measurement pass at all.
 */

import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import {
	BUNDLED_CLI_PATH,
	convertDocx,
	paginateCheck,
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
const parity = fixture("parity-sections.docx");
const plain = fixture("single-page.docx");

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

const results = [];
const scratch = await mkdtemp(join(tmpdir(), "dxpdf-paginate-test-"));

console.log("dxpdf Word-parity pagination");

results.push(
	await check("the check finds no parity break in an ordinary document", async () => {
		const report = await paginateCheck({
			input: plain,
			python: PYTHON,
			timeoutMs: 60_000,
		});
		assert.deepEqual(report.boundaries, []);
		return "0 boundaries, so the measurement pass is skipped";
	}),
);

results.push(
	await check("the check finds the odd-page break in the fixture", async () => {
		const report = await paginateCheck({
			input: parity,
			python: PYTHON,
			timeoutMs: 60_000,
		});
		assert.equal(report.boundaries.length, 1);
		assert.equal(report.boundaries[0].type, "oddPage");
		return `${report.boundaries.length} boundary, type ${report.boundaries[0].type}`;
	}),
);

results.push(
	await check("the fixture really is two pages before any filler", async () => {
		const out = join(scratch, "raw.pdf");
		const facts = await convertDocx({
			input: parity,
			output: out,
			env: PYTHON_ENV,
			matchWordPagination: false,
		});
		assert.equal(facts.paginated, false);
		assert.equal(facts.pages, 2, `expected 2 pages, got ${facts.pages}`);
		return `${facts.pages} pages with pagination off`;
	}),
);

results.push(
	await check("convertDocx restores the blank page Word inserts", async () => {
		const out = join(scratch, "paged.pdf");
		const facts = await convertDocx({ input: parity, output: out, env: PYTHON_ENV });
		assert.equal(facts.paginated, true, "expected the pagination pass to run");
		assert.equal(facts.fillers, 1, `expected 1 filler, got ${facts.fillers}`);
		assert.equal(facts.boundaries, 1);
		assert.equal(facts.pages, 3, `expected 3 pages, got ${facts.pages}`);
		assert.equal(
			(await readFile(out)).subarray(0, 5).toString("latin1"),
			"%PDF-",
		);
		return `2 -> ${facts.pages} pages, ${facts.fillers} blank page restored`;
	}),
);

results.push(
	await check("an ordinary document reports no pagination work", async () => {
		const out = join(scratch, "plain.pdf");
		const facts = await convertDocx({ input: plain, output: out, env: PYTHON_ENV });
		assert.equal(facts.paginated, false);
		assert.equal(facts.fillers, 0);
		assert.equal(facts.boundaries, 0);
		assert.equal(facts.paginationSkipped, undefined, "nothing should have been skipped");
		return "1 page, no pass, nothing skipped";
	}),
);

results.push(
	await check("maxPaginateSections refuses the pass and says why", async () => {
		const out = join(scratch, "capped.pdf");
		const facts = await convertDocx({
			input: parity,
			output: out,
			env: PYTHON_ENV,
			maxPaginateSections: 0,
		});
		assert.equal(facts.paginated, false);
		assert.equal(facts.pages, 2, "the unconverged document must still convert");
		assert.match(facts.paginationSkipped ?? "", /maxPaginateSections/);
		return facts.paginationSkipped;
	}),
);

await rm(scratch, { recursive: true, force: true });

const passed = results.filter(Boolean).length;
console.log(`\n${passed}/${results.length} checks passed`);
process.exit(passed === results.length ? 0 : 1);
