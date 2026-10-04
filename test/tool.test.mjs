/**
 * Tool-contract checks: load the real plugin module, capture the definition it
 * registers, and run that definition's `execute` against real files.
 *
 * Run through the stub hook so `@deepseek-ai/dsh-tools` resolves:
 *   node --import ./test/hooks.mjs test/tool.test.mjs [sample.docx]
 *
 * The argument defaults to the fixture committed beside this file.
 */

import assert from "node:assert/strict";
import { copyFile, mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const sourceDocx =
	process.argv[2] ??
	fileURLToPath(new URL("./fixtures/single-page.docx", import.meta.url));

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
const scratch = await mkdtemp(join(tmpdir(), "dxpdf-tool-test-"));

console.log("dxpdf plugin tool contract");

const module = await import(new URL("../lib/index.js", import.meta.url).href);

results.push(
	await check("exports the plugin identity the loader row expects", async () => {
		assert.equal(module.name, "tool-dxpdf");
		assert.deepEqual(module.inject, ["tools", "systemPrompt"]);
		assert.equal(typeof module.apply, "function");
		return `name=${module.name}, inject=[${module.inject.join(", ")}]`;
	}),
);

const registered = [];
const sections = [];
module.apply(
	{
		tools: { register: (definition) => registered.push(definition) },
		systemPrompt: { section: (section) => sections.push(section) },
	},
	undefined,
);

results.push(
	await check("registers the standing DOCX-to-PDF prompt policy", async () => {
		assert.equal(sections.length, 1, `expected one section, got ${sections.length}`);
		const section = sections[0];
		assert.equal(section.name, module.POLICY_SECTION);
		assert.ok(Number.isFinite(section.order), "order must be a finite number");
		assert.match(section.text, /dxpdf_convert/);
		assert.match(section.text, /LibreOffice/);
		assert.match(section.text, /XLSX and PPTX\s+conversions are unaffected/);
		return `${section.name} at order ${section.order}`;
	}),
);

results.push(
	await check("the policy can be switched off by configuration", async () => {
		const kept = [];
		module.apply(
			{
				tools: { register: () => {} },
				systemPrompt: { section: (section) => kept.push(section) },
			},
			{ announcePolicy: false },
		);
		assert.equal(kept.length, 0, "announcePolicy: false must register no section");
		assert.equal(sections.length, 1, "the first registration must be untouched");
		return "announcePolicy: false registers nothing";
	}),
);

results.push(
	await check("registers exactly one tool named dxpdf_convert", async () => {
		assert.equal(registered.length, 1);
		assert.equal(registered[0].name, "dxpdf_convert");
		assert.equal(typeof registered[0].execute, "function");
		return registered[0].name;
	}),
);

const tool = registered[0];

results.push(
	await check("declares input as required and output/image_dpi as optional", async () => {
		const params = tool.parameters;
		assert.deepEqual(Object.keys(params).sort(), ["image_dpi", "input", "output"]);
		assert.equal(params.input.required, true);
		assert.equal(params.input.type, "string");
		assert.equal(params.output.required, undefined);
		assert.equal(params.image_dpi.type, "number");
		return `params: ${Object.keys(params).join(", ")}`;
	}),
);

results.push(
	await check("declares an object-rooted output schema with a render", async () => {
		const schema = tool.output.schema;
		assert.equal(schema.type, "object");
		assert.equal(schema.additionalProperties, false);
		assert.equal(typeof tool.output.render, "function");
		for (const key of ["input", "output", "imageDpi", "bytes", "elapsedMs", "runner", "normalized", "paginated"]) {
			assert.equal(schema.properties[key]?.required, true, `${key} should be required`);
		}
		assert.equal(schema.properties.pages.required, undefined, "pages should be optional");
		return `required: ${Object.entries(schema.properties).filter(([, v]) => v.required).map(([k]) => k).join(", ")}`;
	}),
);

const workspace = await mkdtemp(join(tmpdir(), "dxpdf-tool-cwd-"));
await copyFile(sourceDocx, join(workspace, "relative.docx"));
const signal = new AbortController().signal;

results.push(
	await check("executes with a workspace-relative input and default output", async () => {
		const value = await tool.execute(
			{ input: "relative.docx" },
			{ signal, agent: { session: { header: { cwd: workspace } } } },
		);
		assert.equal(value.input, join(workspace, "relative.docx"));
		assert.equal(value.output, join(workspace, "relative.pdf"));
		assert.ok(value.bytes > 1000, `expected a real PDF, got ${value.bytes} bytes`);
		assert.equal(value.imageDpi, 220);
		assert.equal(value.pages, 1);
		assert.equal((await readFile(value.output)).subarray(0, 5).toString("latin1"), "%PDF-");
		return `${value.bytes} bytes, ${value.pages} page, ${value.elapsedMs} ms, runner=${value.runner}`;
	}),
);

results.push(
	await check("renders the declared output to model-facing text", async () => {
		const value = await tool.execute(
			{ input: "relative.docx", output: "rendered.pdf", image_dpi: 300 },
			{ signal, agent: { session: { header: { cwd: workspace } } } },
		);
		const content = tool.output.render({}, value);
		assert.equal(content.length, 1);
		assert.equal(content[0].type, "text");
		assert.match(content[0].text, /Converted .* -> .* \(.*, \d+ ms, image_dpi=300, 1 page\)\./);
		return content[0].text;
	}),
);

results.push(
	await check("omits pages when the count is unknown but keeps other fields", async () => {
		const value = await tool.execute(
			{ input: "relative.docx", output: "no-pages.pdf" },
			{ signal, agent: { session: { header: { cwd: workspace } } } },
		);
		assert.equal(typeof value.bytes, "number");
		assert.equal(typeof value.runner, "string");
		// pages must be absent or an integer — never null/undefined-valued.
		assert.ok(value.pages === undefined || Number.isInteger(value.pages));
		const text = tool.output.render({}, { ...value, pages: undefined })[0].text;
		assert.doesNotMatch(text, /undefined/);
		return "clean when pages is absent";
	}),
);

results.push(
	await check("rejects a PDF input with a pointed message", async () => {
		await assert.rejects(
			() =>
				tool.execute(
					{ input: "relative.pdf" },
					{ signal, agent: { session: { header: { cwd: workspace } } } },
				),
			/reads DOCX, not PDF/,
		);
		return "threw as expected";
	}),
);

results.push(
	await check("rejects an output that would overwrite the input", async () => {
		await assert.rejects(
			() =>
				tool.execute(
					{ input: "relative.docx", output: "relative.docx" },
					{ signal, agent: { session: { header: { cwd: workspace } } } },
				),
			/would overwrite the input/,
		);
		return "threw as expected";
	}),
);

results.push(
	await check("falls back to process.cwd() when there is no agent session", async () => {
		const previous = process.cwd();
		process.chdir(workspace);
		try {
			const value = await tool.execute({ input: "relative.docx", output: "nocwd.pdf" }, { signal });
			assert.equal(value.input, join(workspace, "relative.docx"));
			return "resolved against process.cwd()";
		} finally {
			process.chdir(previous);
		}
	}),
);

results.push(
	await check("presentCall describes the call without running it", async () => {
		const card = tool.presentCall({ input: "a.docx", image_dpi: 96 });
		assert.equal(card.card, "generic");
		assert.equal(card.title, "Convert DOCX to PDF");
		assert.match(card.content[0].text, /a\.docx -> a\.pdf \(image_dpi=96\)/);
		return card.content[0].text;
	}),
);

await rm(scratch, { recursive: true, force: true });
await rm(workspace, { recursive: true, force: true });

const passed = results.filter(Boolean).length;
console.log(`\n${passed}/${results.length} checks passed`);
process.exit(passed === results.length ? 0 : 1);
