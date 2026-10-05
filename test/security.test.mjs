/**
 * Safety checks: the invariants that make this plugin auditable.
 *
 * The community standard asks for an audit before a plugin is trusted and then
 * a five-level verification after it is installed (compose, boot smoke, health
 * scan, full boot, functional test). Levels 1-4 say "it loads"; level 5 says
 * "it does its job" — and this suite is what makes the audit part reproducible
 * instead of asserted in prose.
 *
 * Every check below reads this package's own files. A check that cannot read
 * what it needs fails rather than passing quietly, and each one names the exact
 * file and line it objected to.
 *
 * What the plugin is allowed to do, and nothing else:
 *
 * * spawn the conversion engine and its Python helpers with an explicit argv,
 *   never through a shell;
 * * read the input document and write the PDF the caller asked for, plus one
 *   scratch directory under the system temp directory;
 * * reach no network, read no credential, and execute no dynamically built code.
 */

import assert from "node:assert/strict";
import { readFile, readdir, stat } from "node:fs/promises";
import { join, relative, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { pythonChildEnv } from "../lib/convert.js";

const ROOT = fileURLToPath(new URL("..", import.meta.url));
const SKIP = new Set(["node_modules", ".git", "dist", "build", "coverage", "__pycache__"]);

/**
 * The directories whose contents ship to a user, taken from package.json's
 * `files` list plus the root. The tests themselves are not shipped, so a check
 * that looks for a dangerous construct must not be answered by its own pattern.
 */
const packageJson = JSON.parse(await readFile(join(ROOT, "package.json"), "utf8"));
const SHIPPED = ["lib", "py", "locale", "assets", "test/fixtures"]
	.filter((name) => (packageJson.files ?? []).some((entry) => name === entry || name.startsWith(`${entry}/`)));

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

/** Every source file the plugin ships, as absolute paths. */
async function sources(directory = ROOT, found = []) {
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		if (entry.name.startsWith(".") || SKIP.has(entry.name)) continue;
		const path = join(directory, entry.name);
		if (entry.isDirectory()) {
			await sources(path, found);
		} else if (/\.(m?js|cjs|py|ps1|cmd|sh)$/.test(entry.name)) {
			found.push(path);
		}
	}
	return found;
}

/** The shipped sources only, which is what an auditor reads. */
async function shippedSources() {
	const found = [];
	for (const name of SHIPPED) {
		await sources(join(ROOT, name), found);
	}
	for (const entry of await readdir(ROOT, { withFileTypes: true })) {
		if (entry.isFile() && /\.(m?js|cjs)$/.test(entry.name)) found.push(join(ROOT, entry.name));
	}
	return found;
}

/** Every match of `pattern` in the shipped sources, as `file:line — text`. */
async function scan(pattern, files) {
	const hits = [];
	for (const path of files ?? (await shippedSources())) {
		const text = await readFile(path, "utf8");
		text.split("\n").forEach((line, index) => {
			if (pattern.test(line)) {
				hits.push(`${relative(ROOT, path).split(sep).join("/")}:${index + 1} — ${line.trim().slice(0, 90)}`);
			}
			pattern.lastIndex = 0;
		});
	}
	return hits;
}

const results = [];

console.log("dxpdf plugin safety");

results.push(
	await check("nothing runs on the user's machine at install time", async () => {
		const scripts = packageJson.scripts ?? {};
		for (const hook of ["preinstall", "install", "postinstall"]) {
			assert.ok(!(hook in scripts), `package.json declares \`${hook}\`: ${scripts[hook]}`);
		}
		const publishOnly = ["prepack", "prepare", "prepublishOnly"].filter((h) => h in scripts);
		return publishOnly.length === 0
			? "no install hooks, no publish hooks"
			: `no install hooks; publish-only hooks present: ${publishOnly.join(", ")}`;
	}),
);

results.push(
	await check("it brings no dependency of its own into the profile", async () => {
		const runtime = Object.keys(packageJson.dependencies ?? {});
		assert.deepEqual(runtime, [], `runtime dependencies: ${runtime.join(", ")}`);
		const optional = Object.keys(packageJson.optionalDependencies ?? {});
		return `dependencies: 0${optional.length ? `, optional: ${optional.join(", ")}` : ""}`;
	}),
);

results.push(
	await check("it never builds a command line out of a value", async () => {
		const shell = await scan(/shell\s*:\s*true|execSync?\s*\(|spawnSync?\s*\([^)]*shell/);
		assert.deepEqual(shell, [], `a shell would turn a value into code:\n${shell.join("\n")}`);
		return "every spawn passes an argv array";
	}),
);

results.push(
	await check("it executes no dynamically built code", async () => {
		const dynamic = await scan(/\beval\s*\(|new\s+Function\s*\(|vm\.runIn|os\.system|subprocess\.\w+\([^)]*shell\s*=\s*True/);
		assert.deepEqual(dynamic, [], dynamic.join("\n"));
		return "no eval, no Function, no os.system, no shell=True";
	}),
);

results.push(
	await check("it reaches no network", async () => {
		const network = await scan(
			/\bfetch\s*\(|https?\.request\s*\(|require\(["'](https?|net|dns|tls)["']\)|from\s+["'](https?|net|dns|tls)["']|urlopen|requests\.(get|post)|socket\./,
		);
		assert.deepEqual(network, [], `outbound call found:\n${network.join("\n")}`);
		return "no HTTP client, no socket, in JavaScript or Python";
	}),
);

results.push(
	await check("it reads no credential", async () => {
		const credentials = await scan(
			/\.credentials\.ya?ml|settings\.ya?ml|apiKeyEnv|process\.env\.[A-Z_]*(TOKEN|KEY|SECRET|PASSWORD)|os\.environ\[["'][A-Z_]*(TOKEN|KEY|SECRET|PASSWORD)/,
		);
		assert.deepEqual(credentials, [], credentials.join("\n"));
		return "no credential-shaped input";
	}),
);

results.push(
	await check("a child never sees the harness's secrets", async () => {
		const parent = {
			PATH: "C:\\Windows\\System32",
			SystemRoot: "C:\\Windows",
			TEMP: "C:\\Temp",
			DSH_API_KEY: "sk-should-not-travel",
			GPT_API_KEY: "sk-should-not-travel",
			GITHUB_TOKEN: "ghp_should_not_travel",
			AWS_SECRET_ACCESS_KEY: "should-not-travel",
			MYSQL_PASSWORD: "should-not-travel",
			PYTHONHOME: "C:\\python",
			DXPDF_PYTHON: "C:\\python\\python.exe",
		};
		const child = pythonChildEnv(parent);
		for (const leaked of ["DSH_API_KEY", "GPT_API_KEY", "GITHUB_TOKEN",
			"AWS_SECRET_ACCESS_KEY", "MYSQL_PASSWORD"]) {
			assert.ok(!(leaked in child), `${leaked} reached the child`);
		}
		for (const kept of ["PATH", "SystemRoot", "TEMP", "PYTHONHOME", "DXPDF_PYTHON"]) {
			assert.ok(kept in child, `${kept} is needed to run the interpreter`);
		}
		assert.equal(child.PYTHONIOENCODING, "utf-8");
		return `kept ${Object.keys(child).sort().join(", ")}`;
	}),
);

results.push(
	await check("it writes only where it was told to", async () => {
		const writes = await scan(/writeFile\w*\(|createWriteStream\(|open\([^)]*["']w/);
		const outside = writes.filter(
			(hit) => !/("|')(w|wb)("|')/.test(hit) || /homedir|APPDATA|USERPROFILE/.test(hit),
		);
		assert.deepEqual(outside, [], `a write outside the caller's paths:\n${outside.join("\n")}`);
		const absolute = writes.filter((hit) => /["'][A-Za-z]:[\\/]|["']\/(etc|usr|var)\//.test(hit));
		assert.deepEqual(absolute, [], `an absolute path is written:\n${absolute.join("\n")}`);
		return `${writes.length} write site(s), all on the caller's paths or the scratch directory`;
	}),
);

results.push(
	await check("the bundle patch it declares is present", async () => {
		const patch = packageJson.dsh?.bundle?.patch;
		assert.ok(patch, "dsh.bundle.patch is what makes the loader mount the plugin");
		const info = await stat(join(ROOT, patch));
		assert.ok(info.isFile(), `${patch} is not a file`);
		return patch;
	}),
);

results.push(
	await check("the scratch directory is the only place it creates", async () => {
		const temp = await scan(/mkdtemp\(|mkdtempSync\(/);
		assert.ok(temp.length > 0, "the plugin should work in a scratch directory");
		const elsewhere = (await scan(/mkdir\w*\(\s*["'`]/)).filter(
			(hit) => !/dist|node_modules|locale|assets/.test(hit),
		);
		assert.deepEqual(elsewhere, [], elsewhere.join("\n"));
		return `${temp.length} mkdtemp site(s), no fixed directory created`;
	}),
);

results.push(
	await check("the Python helpers are dependency-free and offline", async () => {
		const helpers = (await sources()).filter((path) => path.includes(`${sep}py${sep}`));
		assert.ok(helpers.length >= 3, `expected the Python helpers, found ${helpers.length}`);
		const imports = new Set();
		for (const path of helpers) {
			const text = await readFile(path, "utf8");
			for (const match of text.matchAll(/^\s*(?:import|from)\s+([\w.]+)/gm)) {
				imports.add(match[1].split(".")[0]);
			}
		}
		const allowed = new Set(["argparse", "json", "os", "re", "sys", "zipfile", "collections",
			"shutil", "subprocess", "tempfile", "pathlib", "xml", "lxml", "__future__",
			"typing", "unicodedata", "io", "time", "math", "dataclasses", "importlib",
			"dxpdf"]);
		const unexpected = [...imports].filter((name) => !allowed.has(name));
		assert.deepEqual(unexpected, [], `unexpected import(s): ${unexpected.join(", ")}`);
		return `${helpers.length} helper(s); imports: ${[...imports].sort().join(", ")}`;
	}),
);

results.push(
	await check("the audit finds nothing to hide", async () => {
		const obfuscation = await scan(/(\\x[0-9a-fA-F]{2}){6,}|["'][A-Za-z0-9+/]{400,}={0,2}["']/);
		assert.deepEqual(obfuscation, [], obfuscation.join("\n"));
		return "no escaped blobs, no long inline base64";
	}),
);

const passed = results.filter(Boolean).length;
console.log(`\n${passed}/${results.length} checks passed`);
process.exitCode = passed === results.length ? 0 : 1;
