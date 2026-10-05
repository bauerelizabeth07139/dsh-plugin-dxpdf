/**
 * Stand-in for `@deepseek-ai/dsh-tools` used by `hooks.mjs`.
 *
 * The real `defineTool` is not an identity function: it compiles the tool's
 * schemas and throws when one of them is outside the dialect the host accepts.
 * A stub that returned the definition verbatim therefore passed while the host
 * refused the plugin — which is exactly how an output schema missing
 * `additionalProperties` on an object reached a running profile and stopped the
 * tool from registering at all.
 *
 * So this stub checks the rule that bit us, and names the path the way the host
 * does, so a future failure reads the same in a test as it did at boot:
 *
 *     JsonSchemaError: unsupported JSON schema:
 *     schema.properties.mathConstructs.additionalProperties must be explicitly true or false
 *
 * The dialect is the one `defineTool` accepts: `type` from the JSON Schema
 * primitives, `properties` for an object, `items` for an array, `oneOf` for a
 * union, and `additionalProperties` stated on every object.
 */

/** The marker the host uses for a schema it will not compile. */
export class JsonSchemaError extends Error {
	constructor(message) {
		super(message);
		this.name = "JsonSchemaError";
	}
}

const TYPES = new Set(["string", "number", "integer", "boolean", "object", "array", "null"]);

/**
 * Walk one schema node and refuse what the host refuses.
 * @param {unknown} node - the schema node.
 * @param {string} path - its path, for the message.
 * @returns {void}
 */
function checkSchema(node, path) {
	if (node === null || typeof node !== "object") return;
	const schema = /** @type {Record<string, unknown>} */ (node);

	if (schema.type !== undefined && !TYPES.has(String(schema.type))) {
		throw new JsonSchemaError(`unsupported JSON schema: schema${path}.type ${JSON.stringify(schema.type)} is not a JSON Schema type`);
	}
	if (schema.type === "object") {
		if (typeof schema.additionalProperties !== "boolean") {
			throw new JsonSchemaError(
				`unsupported JSON schema: schema${path}.additionalProperties must be explicitly true or false`,
			);
		}
		for (const [key, value] of Object.entries(schema.properties ?? {})) {
			checkSchema(value, `${path}.properties.${key}`);
		}
	}
	if (Array.isArray(schema.oneOf)) {
		schema.oneOf.forEach((value, index) => checkSchema(value, `${path}.oneOf[${index}]`));
	}
	if (schema.items !== undefined) checkSchema(schema.items, `${path}.items`);
}

/**
 * Register a tool definition, checking its schemas first.
 * @param {Record<string, any>} definition - the definition the plugin passes.
 * @returns {Record<string, any>} the same definition.
 */
export function defineTool(definition) {
	if (definition?.output?.schema !== undefined) {
		checkSchema(definition.output.schema, "");
	}
	for (const [name, value] of Object.entries(definition?.parameters ?? {})) {
		checkSchema(value, `.parameters.${name}`);
	}
	return definition;
}

/** The real package throws this marker for a cancelled tool call. */
export const TOOL_ABORTED = "TOOL_ABORTED";
