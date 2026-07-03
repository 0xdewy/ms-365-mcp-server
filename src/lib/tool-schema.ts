import { z } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';
import type { api } from '../generated/client.js';

type ToolEndpoint = (typeof api.endpoints)[number];
type ParameterLocation = 'Path' | 'Query' | 'Body' | 'Header';

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
 */
export function describeToolSchema(
  tool: ToolEndpoint,
  llmTip: string | undefined,
  descriptionOverride?: string
): {
  name: string;
  method: string;
  path: string;
  description: string;
  llmTip?: string;
  parameters: Array<{
    name: string;
    in: ParameterLocation;
    required: boolean;
    description?: string;
    schema: unknown;
  }>;
} {
  const params = (tool.parameters ?? []).map((p) => {
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

  return {
    name: tool.alias,
    method: tool.method.toUpperCase(),
    path: tool.path,
    description: descriptionOverride ?? tool.description ?? '',
    ...(llmTip ? { llmTip } : {}),
    parameters: params,
  };
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
