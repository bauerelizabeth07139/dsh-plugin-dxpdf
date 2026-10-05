/**
 * `dxpdf_convert` — a model-facing tool that converts a Microsoft Word DOCX
 * file to PDF with the `dxpdf` Rust/Skia engine, without Office or LibreOffice.
 *
 * The plugin is a thin, typed shell around `./convert.js`: it owns argument
 * validation, path resolution against the session working directory, caller
 * cancellation, and the rendered result. All subprocess mechanics — runner
 * discovery from the `DXPDF_*` machine variables, argv construction, timeout,
 * page counting, and the transparent OOXML repair pass — live in the core
 * module so they can be tested headlessly.
 *
 * @module dsh-plugin-dxpdf
 */

import { dirname } from "node:path";
import { defineTool } from "@deepseek-ai/dsh-tools";
import {
	convertDocx,
	DEFAULT_IMAGE_DPI,
	DEFAULT_MAX_PAGINATE_SECTIONS,
	DEFAULT_PAGINATE_TIMEOUT_MS,
	DEFAULT_TIMEOUT_MS,
	IMAGE_DPI_RANGE,
	resolvePath,
} from "./convert.js";

/** Plugin identity used by the loader row. */
export const name = "tool-dxpdf";
/** The registries this plugin registers into. */
export const inject = ["tools", "systemPrompt"];

const [MIN_DPI, MAX_DPI] = IMAGE_DPI_RANGE;

/**
 * Where the DOCX-to-PDF policy sits in the assembled prompt: after the write
 * tool's section (1200) and before the edit tool's (1300), so it reads as part
 * of the tool guidance rather than as an afterthought.
 */
const POLICY_ORDER = 1250;

/** Stable section name, so a later composition can target or replace it. */
export const POLICY_SECTION = "dxpdf:docx-to-pdf-policy";

/**
 * The standing instruction that makes this the deployment's DOCX-to-PDF
 * converter. A tool description only competes with other tools; a system
 * section is what settles the choice against the bundled office-docx skill,
 * whose LibreOffice `convert` step would otherwise win for a .docx.
 */
const POLICY = `## DOCX to PDF

Convert every DOCX-to-PDF request with the \`dxpdf_convert\` tool. It is this
deployment's converter: standalone (no Microsoft Office, no LibreOffice, no
network), it repairs the schema-invalid constructs WPS Office writes, and it
reports the resulting page count, size and duration. Do not use bundled
LibreOffice, \`soffice\`, or a Python rendering library to turn a .docx into a
.pdf. The office-docx skill still governs creating, editing and structurally
checking Word documents. Fall back to that skill's bundled-LibreOffice path
only when \`dxpdf_convert\` reports the document exceeds the engine's coverage
(charts, SmartArt, tracked changes), and say so when you do. XLSX and PPTX
conversions are unaffected by this policy.`;

const DESCRIPTION =
	"Convert a Microsoft Word DOCX file to PDF with the dxpdf engine (Rust + Skia) — " +
	"no Microsoft Office, no LibreOffice, no network. Use it whenever a .docx must " +
	"become a .pdf. Input must be an existing .docx; the tool writes a .pdf and " +
	"reports its size, the conversion time and the page count. An existing output " +
	"file is overwritten. Pass an absolute path, or a path relative to the session " +
	"working directory. DOCX only: .doc, .odt, .rtf and .pdf inputs are not supported. " +
	"Documents written by WPS Office carry a few schema-invalid constructs that " +
	"Word tolerates but dxpdf rejects; the tool repairs those automatically and " +
	"converts a corrected copy, leaving the source file untouched, so no " +
	"preparation is needed. This is the deployment's DOCX-to-PDF converter: reach " +
	"for it rather than LibreOffice or a Python rendering library whenever a Word " +
	"document must become a PDF.";

/**
 * Register the `dxpdf_convert` tool.
 * @param {object} ctx - registrant context; must carry `tools` and `systemPrompt`.
 * @param {object} [config] - optional deployment configuration.
 * @param {number} [config.timeoutMs] - wall-clock budget per conversion.
 * @param {string} [config.pythonPath] - explicit interpreter override.
 * @param {string} [config.cliPath] - explicit CLI-script override.
 * @param {string} [config.executable] - explicit `dxpdf` shim override.
 * @param {"first" | "last"} [config.duplicatePolicy] - which occurrence of a
 *   duplicated property element a repair keeps. Defaults to `first`.
 * @param {boolean} [config.announcePolicy] - whether to register the standing
 *   DOCX-to-PDF system-prompt policy. Defaults to `true`.
 * @param {boolean} [config.matchWordPagination] - converge odd/even section
 *   breaks onto Word's pagination. Defaults to `true`.
 * @param {number} [config.paginateTimeoutMs] - budget for the pagination pass.
 * @param {number} [config.maxPaginateSections] - refuse the pagination pass
 *   beyond this many parity breaks.
 */
export function apply(ctx, config) {
	const settings = config ?? {};
	const timeoutMs =
		Number.isFinite(settings.timeoutMs) && settings.timeoutMs > 0
			? settings.timeoutMs
			: DEFAULT_TIMEOUT_MS;
	const paginateTimeoutMs =
		Number.isFinite(settings.paginateTimeoutMs) && settings.paginateTimeoutMs > 0
			? settings.paginateTimeoutMs
			: DEFAULT_PAGINATE_TIMEOUT_MS;
	const maxPaginateSections =
		Number.isFinite(settings.maxPaginateSections) && settings.maxPaginateSections >= 0
			? settings.maxPaginateSections
			: DEFAULT_MAX_PAGINATE_SECTIONS;
	const matchWordPagination = settings.matchWordPagination !== false;
	const policy = settings.duplicatePolicy === "last" ? "last" : "first";
	const announcePolicy = settings.announcePolicy !== false;

	if (announcePolicy && ctx.systemPrompt !== undefined) {
		// `section()` returns its own Cordis effect disposer, so the section
		// disappears with the plugin rather than accumulating across reloads.
		ctx.systemPrompt.section({
			name: POLICY_SECTION,
			order: POLICY_ORDER,
			text: POLICY,
		});
	}

	ctx.tools.register(
		defineTool({
			name: "dxpdf_convert",
			description: DESCRIPTION,
			parameters: {
				input: {
					type: "string",
					required: true,
					description:
						"Path to the source .docx file. Absolute, or relative to the session working directory.",
				},
				output: {
					type: "string",
					description:
						"Path for the produced .pdf. Defaults to the input path with its extension replaced by .pdf. An existing file there is overwritten.",
				},
				image_dpi: {
					type: "number",
					description: `Resolution embedded raster images are downsampled to, ${MIN_DPI}-${MAX_DPI}. Defaults to ${DEFAULT_IMAGE_DPI} (what Word uses); raise it (e.g. 300) for print quality, lower it (e.g. 96) for smaller files. Images are never upsampled past their source resolution.`,
				},
			},
			output: {
				schema: {
					type: "object",
					additionalProperties: false,
					properties: {
						input: {
							type: "string",
							required: true,
							description: "Absolute path of the DOCX that was read.",
						},
						output: {
							type: "string",
							required: true,
							description: "Absolute path of the PDF that was written.",
						},
						imageDpi: {
							type: "number",
							required: true,
							description: "Embedded-image resolution actually used.",
						},
						bytes: {
							type: "integer",
							required: true,
							description: "Size of the produced PDF in bytes.",
						},
						elapsedMs: {
							type: "integer",
							required: true,
							description: "Wall-clock milliseconds spent in the conversion.",
						},
						pages: {
							type: "integer",
							description:
								"Page count read back from the produced PDF. Omitted when the pages could not be counted.",
						},
						runner: {
							type: "string",
							required: true,
							description:
								"How the dxpdf engine was located, for troubleshooting a machine's setup.",
						},
						normalized: {
							type: "boolean",
							required: true,
							description:
								"Whether the document had to be repaired before dxpdf would accept it. True means its XML was not schema-valid (WPS Office writes such files) and a corrected copy was converted instead; the source file was not modified.",
						},
						repairs: {
							type: "integer",
							description:
								"Number of schema violations repaired. Present only when `normalized` is true.",
						},
						paginated: {
							type: "boolean",
							required: true,
							description:
								"Whether blank pages were inserted so odd/even section breaks land on the page parity Word promises. dxpdf treats those breaks as a plain nextPage, so without this pass the blank pages are missing.",
						},
						fillers: {
							type: "integer",
							description:
								"Number of blank pages inserted. Present only when `paginated` is true.",
						},
						boundaries: {
							type: "integer",
							description:
								"Number of odd/even section breaks found. Present only when `paginated` is true.",
						},
						paginationSkipped: {
							type: "string",
							description:
								"Why Word-parity pagination was not applied, when the document has odd/even section breaks but the pass was refused or failed. Absent when it was applied or was not needed.",
						},
						mathEquations: {
							type: "integer",
							required: true,
							description:
								"How many equations in the document use a construct dxpdf does not lay out.",
						},
						mathLowered: {
							type: "integer",
							required: true,
							description:
								"How many equations were rewritten into the subset dxpdf renders. Without this the engine reports success and leaves the formula out of the page.",
						},
						mathConstructs: {
							type: "object",
							required: true,
							// The keys are construct names, so the map is open by
							// definition — and the host's schema compiler refuses an
							// object schema that leaves this unstated.
							additionalProperties: true,
							description:
								"Which OMML constructs were rewritten, by name (`sSub`, `rad`, `d`, `nary`, `acc`, …).",
						},
						mathSkipped: {
							type: "string",
							description:
								"Why the math pass did not run, when it failed. Absent when it ran or when the document needed nothing.",
						},
					},
				},
				render: (_args, value) => [
					{
						type: "text",
						text:
							`Converted ${value.input} -> ${value.output} ` +
							`(${formatBytes(value.bytes)}, ${value.elapsedMs} ms, ` +
							`image_dpi=${value.imageDpi}${value.pages === undefined ? "" : `, ${value.pages} page${value.pages === 1 ? "" : "s"}`}).` +
							`${value.normalized ? ` Repaired ${value.repairs} schema violation(s) in a copy first; the source file was not modified.` : ""}` +
							`${value.mathLowered > 0 ? ` Lowered ${value.mathLowered} equation(s) dxpdf would have dropped (${Object.entries(value.mathConstructs ?? {}).slice(0, 6).map(([name, count]) => `${name}×${count}`).join(", ")}) in a copy first; the source file was not modified.` : ""}` +
							`${value.mathSkipped === undefined ? "" : ` The math pass did not run: ${value.mathSkipped}.`}` +
							`${value.paginated ? ` Restored ${value.fillers} blank page(s) so the ${value.boundaries} odd/even section break(s) land on the page parity Word promises.` : ""}` +
							`${value.paginationSkipped === undefined ? "" : ` Word-parity pagination was not applied: ${value.paginationSkipped}.`}`,
					},
				],
			},
			async execute(args, exec) {
				const cwd = exec.agent?.session.header.cwd ?? process.cwd();
				const input = resolvePath(args.input, cwd);
				const output =
					args.output === undefined
						? replaceExtension(input, ".pdf")
						: resolvePath(args.output, cwd);

				if (input.toLowerCase().endsWith(".pdf")) {
					throw new Error(
						`dxpdf_convert reads DOCX, not PDF: got ${args.input}. Pass the .docx source file.`,
					);
				}
				if (output === input) {
					throw new Error(
						`invalid output: it would overwrite the input file ${input}. Choose a different output path.`,
					);
				}

				const result = await convertDocx({
					input,
					output,
					imageDpi: args.image_dpi ?? DEFAULT_IMAGE_DPI,
					timeoutMs,
					signal: exec.signal,
					cwd: dirname(input),
					pythonPath: settings.pythonPath,
					cliPath: settings.cliPath,
					executable: settings.executable,
					policy,
					matchWordPagination,
					paginateTimeoutMs,
					maxPaginateSections,
				});

				return {
					input,
					output,
					imageDpi: args.image_dpi ?? DEFAULT_IMAGE_DPI,
					bytes: result.bytes,
					elapsedMs: result.elapsedMs,
					...(result.pages === undefined ? {} : { pages: result.pages }),
					runner: result.runner,
					normalized: result.normalized,
					...(result.normalized ? { repairs: result.edits } : {}),
					mathEquations: result.mathEquations,
					mathLowered: result.mathLowered,
					mathConstructs: result.mathConstructs,
					...(result.mathSkipped === undefined
						? {}
						: { mathSkipped: result.mathSkipped }),
					paginated: result.paginated,
					...(result.paginated
						? { fillers: result.fillers, boundaries: result.boundaries }
						: {}),
					...(result.paginationSkipped === undefined
						? {}
						: { paginationSkipped: result.paginationSkipped }),
				};
			},
			presentCall: (args) => ({
				card: "generic",
				title: "Convert DOCX to PDF",
				kind: "other",
				rawInput: args.input,
				content: [
					{
						type: "text",
						text:
							`${args.input} -> ${args.output ?? replaceExtension(args.input, ".pdf")}` +
							`${args.image_dpi === undefined ? "" : ` (image_dpi=${args.image_dpi})`}`,
					},
				],
			}),
		}),
	);
}

/**
 * Replace a path's extension, appending one when there is none.
 * @param {string} path - the source path.
 * @param {string} extension - extension including the leading dot.
 * @returns {string} the path with the new extension.
 */
function replaceExtension(path, extension) {
	const base = path.slice(0, path.length - extensionOf(path).length);
	return `${base}${extension}`;
}

/**
 * The trailing extension of the final path segment, or an empty string. A
 * leading dot names a hidden file rather than an extension, and both Windows
 * and POSIX separators are honoured because a model may supply either.
 * @param {string} path - the path to inspect.
 * @returns {string} the extension including its dot, when present.
 */
function extensionOf(path) {
	const cut = Math.max(path.lastIndexOf("\\"), path.lastIndexOf("/"));
	const segment = path.slice(cut + 1);
	const dot = segment.lastIndexOf(".");
	return dot > 0 ? segment.slice(dot) : "";
}

/**
 * Render a byte count for the model-facing summary.
 * @param {number} bytes - size in bytes.
 * @returns {string} a compact human-readable size.
 */
function formatBytes(bytes) {
	if (bytes < 1024) return `${bytes} B`;
	if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KiB`;
	return `${(bytes / (1024 * 1024)).toFixed(1)} MiB`;
}
