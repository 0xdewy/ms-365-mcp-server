import { z } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';
import type { api } from '../generated/client.js';
import { isDestructiveOperation, type DestructiveCheckConfig } from './destructive-ops.js';

type ToolEndpoint = (typeof api.endpoints)[number];
type ParameterLocation = 'Path' | 'Query' | 'Body' | 'Header';
type DescribedParameter = {
  name: string;
  in: ParameterLocation;
  required: boolean;
  description?: string;
  schema: unknown;
};

const CONFLICT_BEHAVIOR_SCHEMA = {
  type: 'string',
  enum: ['rename', 'replace', 'fail'],
  description: 'How Graph should handle a name conflict.',
} as const;

const COMPACT_BODY_SCHEMAS: Record<string, unknown> = {
  'create-sharepoint-list': {
    type: 'object',
    required: ['displayName', 'list'],
    additionalProperties: true,
    properties: {
      displayName: { type: 'string', description: 'Visible list name.' },
      description: { type: 'string', description: 'Optional list description.' },
      list: {
        type: 'object',
        required: ['template'],
        additionalProperties: true,
        properties: {
          template: {
            type: 'string',
            enum: [
              'genericList',
              'documentLibrary',
              'tasks',
              'calendar',
              'contacts',
              'links',
              'announcements',
              'survey',
            ],
          },
        },
      },
      columns: {
        type: 'array',
        description: 'Optional initial column definitions.',
        items: {
          type: 'object',
          required: ['name'],
          additionalProperties: true,
          properties: {
            name: { type: 'string' },
            displayName: { type: 'string' },
            description: { type: 'string' },
            text: { type: 'object', additionalProperties: true },
            choice: {
              type: 'object',
              additionalProperties: true,
              properties: {
                choices: { type: 'array', items: { type: 'string' } },
              },
            },
            dateTime: { type: 'object', additionalProperties: true },
            number: { type: 'object', additionalProperties: true },
            boolean: { type: 'object', additionalProperties: true },
            currency: { type: 'object', additionalProperties: true },
            hyperlinkOrPicture: { type: 'object', additionalProperties: true },
            personOrGroup: { type: 'object', additionalProperties: true },
            lookup: { type: 'object', additionalProperties: true },
            calculated: { type: 'object', additionalProperties: true },
          },
        },
      },
    },
  },
  'create-sharepoint-list-item': {
    type: 'object',
    required: ['fields'],
    additionalProperties: true,
    properties: {
      fields: {
        type: 'object',
        description: 'SharePoint column internal names mapped to values.',
        additionalProperties: true,
      },
    },
  },
  'update-sharepoint-list-item': {
    type: 'object',
    required: ['fields'],
    additionalProperties: true,
    properties: {
      fields: {
        type: 'object',
        description: 'Only the SharePoint fields to update, keyed by column internal name.',
        additionalProperties: true,
      },
    },
  },
  'create-onedrive-folder': {
    type: 'object',
    required: ['name', 'folder'],
    additionalProperties: true,
    properties: {
      name: { type: 'string', description: 'Folder name to create.' },
      folder: {
        type: 'object',
        description: 'Must be an empty object for a folder create request.',
        additionalProperties: true,
      },
      '@microsoft.graph.conflictBehavior': CONFLICT_BEHAVIOR_SCHEMA,
    },
  },
  'move-rename-onedrive-item': {
    type: 'object',
    additionalProperties: true,
    properties: {
      name: { type: 'string', description: 'New item name when renaming.' },
      parentReference: {
        type: 'object',
        description: 'Target parent folder reference when moving.',
        additionalProperties: true,
        properties: {
          id: { type: 'string', description: 'Target parent driveItem id.' },
          driveId: { type: 'string', description: 'Target drive id, when moving across drives.' },
        },
      },
      '@microsoft.graph.conflictBehavior': CONFLICT_BEHAVIOR_SCHEMA,
    },
  },
};

/**
 * Subset of EndpointConfig needed to describe a tool's schema in discovery
 * mode. Kept as a structural type so we don't import the full EndpointConfig
 * from graph-tools.ts (which would create a circular dependency).
 */
export interface ToolSchemaConfig extends DestructiveCheckConfig {
  llmTip?: string;
  descriptionOverride?: string;
}

function unwrapOptional(schema: z.ZodTypeAny): { inner: z.ZodTypeAny; optional: boolean } {
  const def = (schema as { _def?: { typeName?: string; innerType?: z.ZodTypeAny } })._def;
  const typeName = def?.typeName;
  if (typeName === 'ZodOptional' || typeName === 'ZodDefault' || typeName === 'ZodNullable') {
    return { inner: def!.innerType!, optional: true };
  }
  return { inner: schema, optional: false };
}

/**
 * Returns a JSON Schema describing every parameter a discovery tool accepts,
 * so an agent can construct a correctly-shaped `parameters` object for execute-tool.
 *
 * Includes synthetic runtime params injected by graph-tools.ts that an agent
 * needs to know about, including pagination/response controls and `confirm`
 * for destructive operations.
 */
export function describeToolSchema(
  tool: ToolEndpoint,
  config: ToolSchemaConfig | undefined
): {
  name: string;
  method: string;
  path: string;
  description: string;
  llmTip?: string;
  parameters: DescribedParameter[];
} {
  const params: DescribedParameter[] = (tool.parameters ?? []).map((p) => {
    const { inner, optional } = unwrapOptional(p.schema as z.ZodTypeAny);
    const isPath = p.type === 'Path';
    const schema =
      p.type === 'Body' && COMPACT_BODY_SCHEMAS[tool.alias]
        ? COMPACT_BODY_SCHEMAS[tool.alias]
        : toJsonSchema(inner);
    return {
      name: p.name,
      in: p.type as ParameterLocation,
      required: isPath || !optional,
      description: p.description,
      schema,
    };
  });

  // Surface the destructive-confirm gate so agents in --discovery mode know
  // to pass `confirm: true`. Without this, every destructive tool returns
  // confirmation_required with no way for the agent to recover from the schema.
  if (isDestructiveOperation(tool.method, config)) {
    params.push({
      name: 'confirm',
      in: 'Query',
      required: false,
      description:
        'For destructive operations when the confirm gate is enabled (MS365_MCP_REQUIRE_CONFIRM=true; off by default). ' +
        'Set to true only after the user has explicitly approved this action. ' +
        'When the gate is on, calls without confirm: true return { error: "confirmation_required" } without touching user data.',
      schema: { type: 'boolean' },
    });
  }

  params.push(...controlParametersFor(tool));

  const llmTip = config?.llmTip;
  return {
    name: tool.alias,
    method: tool.method.toUpperCase(),
    path: tool.path,
    description: config?.descriptionOverride ?? tool.description ?? '',
    ...(llmTip ? { llmTip } : {}),
    parameters: params,
  };
}

function controlParametersFor(tool: ToolEndpoint): DescribedParameter[] {
  const controls: DescribedParameter[] = [
    {
      name: 'includeHeaders',
      in: 'Query',
      required: false,
      description: 'Include response headers such as ETag in the response metadata.',
      schema: { type: 'boolean' },
    },
    {
      name: 'excludeResponse',
      in: 'Query',
      required: false,
      description: 'Return only success or failure instead of the full response body.',
      schema: { type: 'boolean' },
    },
  ];

  if (tool.method.toUpperCase() === 'GET') {
    controls.unshift({
      name: 'fetchAllPages',
      in: 'Query',
      required: false,
      description:
        'Follow @odata.nextLink and merge pages when the server allows pagination. Use with small $top/$select values for bounded exports.',
      schema: { type: 'boolean' },
    });
  }

  return controls.filter(
    (control) => !(tool.parameters ?? []).some((p) => p.name === control.name)
  );
}

function toJsonSchema(schema: z.ZodTypeAny): unknown {
  const jsonSchema = zodToJsonSchema(schema, { target: 'jsonSchema7', $refStrategy: 'seen' });
  const { $schema: _s, ...result } = jsonSchema as Record<string, unknown>;
  return result;
}

interface UtilityDescriptor {
  name: string;
  method: string;
  path: string;
  description: string;
  buildSchema: (ctx: never) => Record<string, z.ZodTypeAny>;
}

// Params reported as `Query` (top-level): execute-tool passes `parameters`
// straight to utility.execute(); `Body` would mislead LLMs into nesting under `body`.
export function describeUtilityToolSchema<C>(
  utility: UtilityDescriptor & { buildSchema: (ctx: C) => Record<string, z.ZodTypeAny> },
  ctx: C
): {
  name: string;
  method: string;
  path: string;
  description: string;
  parameters: Array<{
    name: string;
    in: 'Query';
    required: boolean;
    description?: string;
    schema: unknown;
  }>;
} {
  const schemaMap = utility.buildSchema(ctx);
  const params = Object.entries(schemaMap).map(([name, zodSchema]) => {
    const { inner, optional } = unwrapOptional(zodSchema);
    return {
      name,
      in: 'Query' as const,
      required: !optional,
      description: zodSchema.description,
      schema: toJsonSchema(inner),
    };
  });
  return {
    name: utility.name,
    method: utility.method,
    path: utility.path,
    description: utility.description,
    parameters: params,
  };
}
