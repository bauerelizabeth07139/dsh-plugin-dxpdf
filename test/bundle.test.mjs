/**
 * Static composition check: everything the Harness resolves before it imports
 * a bundle's host code, verified against an installed profile.
 *
 * It answers "will the loader find and mount this bundle" without starting a
 * second Harness against the live profile.
 *
 * It needs a profile that has this bundle installed, and a copy of `js-yaml`
 * (which ships with the Harness, not with this plugin). Point it at both:
 *
 *   DSH_PROFILE=<...>/profiles/desktop DSH_JS_YAML=<...>/js-yaml.mjs \
 *     node test/bundle.test.mjs
 *
 * `DSH_PROFILE` may be either a profile directory or a profile name (the
 * Harness itself exports the name), and `DSH_HOME` is honoured when set. With
 * no profile to inspect it prints SKIP and exits 0, so `npm test` works on a
 * fresh clone.
 */

import assert from "node:assert/strict";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, isAbsolute, join } from "node:path";
import { pathToFileURL } from "node:url";

const requested = process.env.DSH_PROFILE?.trim() ?? "";
const isPath = requested !== "" && (/[/\\]/.test(requested) || isAbsolute(requested));
const profileName = isPath ? process.env.DSH_PROFILE_NAME ?? "desktop" : requested || "desktop";
const PROFILE = isPath
	? requested
	: join(process.env.DSH_HOME?.trim() || join(homedir(), ".dsh"), "profiles", profileName);
const BUNDLE = "dsh-plugin-dxpdf";

if (!existsSync(join(PROFILE, "package.json"))) {
	console.log("dxpdf bundle composition");
	console.log(`  SKIP  no installed profile at ${PROFILE}`);
	console.log("        set DSH_PROFILE to a profile that has this bundle installed");
	process.exit(0);
}

// `js-yaml` is part of the Harness installation, not a dependency of this
// plugin, so allow the caller to point at the copy it wants checked against —
// and look in the neighbouring profiles, where the Harness keeps one.
const yamlCandidates = [
	process.env.DSH_JS_YAML,
	"js-yaml",
	join(PROFILE, "node_modules", "js-yaml", "dist", "js-yaml.mjs"),
];
try {
	for (const sibling of readdirSync(dirname(PROFILE))) {
		yamlCandidates.push(
			join(dirname(PROFILE), sibling, "node_modules", "js-yaml", "dist", "js-yaml.mjs"),
		);
	}
} catch {
	// the profile may be a plain directory rather than one of several
}

let load;
for (const candidate of yamlCandidates) {
	try {
		const specifier = candidate.includes("/") || candidate.includes("\\")
			? pathToFileURL(candidate).href
			: candidate;
		({ load } = await import(specifier));
		if (load) break;
	} catch {
		// try the next candidate
	}
}

if (!load) {
	console.log("dxpdf bundle composition");
	console.log("  SKIP  js-yaml not found — set DSH_JS_YAML to the Harness copy");
	console.log("        (e.g. <profile>/node_modules/js-yaml/dist/js-yaml.mjs)");
	process.exit(0);
}

const results = [];
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

console.log("dxpdf bundle composition");

const manifest = JSON.parse(readFileSync(join(PROFILE, "package.json"), "utf8"));

results.push(
	await check("the profile manifest selects the bundle", async () => {
		const bundles = manifest.dsh?.profile?.bundles ?? [];
		assert.ok(bundles.includes(BUNDLE), `bundles = ${JSON.stringify(bundles)}`);
		return `${bundles.length} bundles, ours last`;
	}),
);

results.push(
	await check("the profile manifest declares the dependency", async () => {
		const spec = manifest.dependencies?.[BUNDLE];
		assert.ok(spec, "dependency missing");
		return spec;
	}),
);

const installed = join(PROFILE, "node_modules", BUNDLE);
const installedManifest = JSON.parse(readFileSync(join(installed, "package.json"), "utf8"));

results.push(
	await check("the installed package name matches the loader row name", async () => {
		assert.equal(installedManifest.name, BUNDLE);
		return installedManifest.name;
	}),
);

results.push(
	await check("the installed package declares a bundle patch", async () => {
		const patch = installedManifest.dsh?.bundle?.patch;
		assert.ok(patch, "dsh.bundle.patch missing");
		assert.ok(existsSync(join(installed, patch)), `${patch} does not exist`);
		return patch;
	}),
);

results.push(
	await check("the installed entry point exists", async () => {
		const main = installedManifest.main ?? "./lib/index.js";
		assert.ok(existsSync(join(installed, main)), `${main} does not exist`);
		return main;
	}),
);

results.push(
	await check("the locale card resolves through the package exports", async () => {
		const exposed = installedManifest.exports?.["./locale/*.json"];
		assert.ok(exposed, "exports['./locale/*.json'] missing");
		for (const language of ["en", "zh"]) {
			const card = JSON.parse(
				readFileSync(join(installed, "locale", `${language}.json`), "utf8"),
			);
			assert.ok(card.meta?.title, `locale/${language}.json meta.title missing`);
			assert.ok(
				card.meta?.description,
				`locale/${language}.json meta.description missing`,
			);
		}
		const icon = installedManifest.icon;
		assert.ok(icon, "icon missing");
		assert.ok(existsSync(join(installed, icon)), `${icon} does not exist`);
		return `${exposed}, icon ${icon}`;
	}),
);

const patchPath = join(installed, installedManifest.dsh.bundle.patch);

results.push(
	await check("the bundle patch parses as YAML", async () => {
		const parsed = load(readFileSync(patchPath, "utf8"));
		assert.ok(Array.isArray(parsed), "a patch layer must be a top-level array");
		return `${parsed.length} entr${parsed.length === 1 ? "y" : "ies"}`;
	}),
);

results.push(
	await check("the bundle patch inserts exactly one loader row naming the package", async () => {
		const parsed = load(readFileSync(patchPath, "utf8"));
		const inserts = parsed.filter((entry) => entry?.insert !== undefined);
		assert.equal(inserts.length, 1, "expected exactly one insert entry");
		const rows = inserts[0].insert;
		assert.ok(Array.isArray(rows) && rows.length === 1, "expected exactly one row");
		assert.equal(rows[0].name, BUNDLE);
		assert.equal(typeof rows[0].id, "string");
		assert.notEqual(rows[0].id.length, 0);
		return `insert -> { id: ${rows[0].id}, name: ${rows[0].name} }`;
	}),
);

results.push(
	await check("no other bundle already claims the same loader row id", async () => {
		const rowId = load(readFileSync(patchPath, "utf8"))[0].insert[0].id;
		const owners = [];
		for (const name of manifest.dsh.profile.bundles) {
			if (name.startsWith("@deepseek-ai/")) continue;
			const dir = join(PROFILE, "node_modules", name);
			const file = join(dir, "package.json");
			if (!existsSync(file)) continue;
			const other = JSON.parse(readFileSync(file, "utf8"));
			const relative = other.dsh?.bundle?.patch;
			if (!relative || !existsSync(join(dir, relative))) continue;
			const entries = load(readFileSync(join(dir, relative), "utf8"));
			for (const entry of entries) {
				for (const row of entry?.insert ?? []) {
					if (row?.id === rowId) owners.push(name);
				}
			}
		}
		assert.deepEqual(owners, [BUNDLE], `row id ${rowId} claimed by ${owners.join(", ")}`);
		return `${rowId} is unique`;
	}),
);

results.push(
	await check("the declared DSH engine range admits the 0.2 line", async () => {
		const range = installedManifest.engines?.dsh;
		assert.ok(range, "engines.dsh missing");
		assert.ok(
			/^>=0\.1\.5-rc\.1 <0\.2\.0-0 \|\| >=0\.2\.0-rc\.0 <0\.3\.0-0$/.test(range),
			`unexpected range ${range}`,
		);
		const runtimeFile =
			process.env.DSH_RUNTIME_JSON ??
			join(PROFILE, "..", "..", "..", "deeph", "resources", "runtime", "primary-runtime", "runtime.json");
		const runtime = existsSync(runtimeFile)
			? JSON.parse(readFileSync(runtimeFile, "utf8")).desktopVersion
			: undefined;
		return runtime === undefined ? range : `${range} admits ${runtime}`;
	}),
);

const passed = results.filter(Boolean).length;
console.log(`\n${passed}/${results.length} checks passed`);
process.exit(passed === results.length ? 0 : 1);
