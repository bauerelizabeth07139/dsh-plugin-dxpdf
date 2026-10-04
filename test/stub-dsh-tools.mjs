/**
 * Stand-in for `@deepseek-ai/dsh-tools` used by `hooks.mjs`.
 *
 * `defineTool` is an identity function in the real package — it exists to give
 * the definition a checked type — so returning the definition verbatim keeps
 * the plugin's registered shape observable to the test.
 */

export function defineTool(definition) {
	return definition;
}

/** The real package throws this marker for a cancelled tool call. */
export const TOOL_ABORTED = "TOOL_ABORTED";
