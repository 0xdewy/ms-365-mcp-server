import { describe, expect, it, vi } from 'vitest';
import { registerGraphTools } from '../src/graph-tools.js';
import { api } from '../src/generated/client.js';
import endpoints from '../src/endpoints.json' with { type: 'json' };

function createMockServer() {
  const tools = new Map<
    string,
    { description: string; schema: Record<string, unknown>; handler: (args: any) => Promise<any> }
  >();

  return {
    tools,
    tool: vi.fn(
      (
        name: string,
        description: string,
        schema: Record<string, unknown>,
        _annotations: Record<string, unknown>,
        handler: (args: any) => Promise<any>
      ) => {
        tools.set(name, { description, schema, handler });
      }
    ),
  };
}

describe('upload-file-content regression', () => {
  it('keeps upload-file-content configured as a binary request', () => {
    const generatedTool = api.endpoints.find(
      (endpoint) => endpoint.alias === 'upload-file-content'
    );
    const sourceConfig = endpoints.find((endpoint) => endpoint.toolName === 'upload-file-content');

    expect(generatedTool?.requestFormat).toBe('binary');
    expect(sourceConfig?.requestFormat).toBe('binary');
  });

  it('sends raw bytes with application/octet-stream for upload-file-content', async () => {
    const graphClient = {
      graphRequest: vi.fn().mockResolvedValue({
        content: [{ type: 'text', text: JSON.stringify({ id: 'item456' }) }],
      }),
    };
    const server = createMockServer();

    registerGraphTools(server as any, graphClient as any, false, '^upload-file-content$');

    await server.tools.get('upload-file-content')!.handler({
      driveId: 'drive123',
      driveItemId: 'item456',
      body: Buffer.from('Hello, world!', 'utf-8').toString('base64'),
    });

    const [path, options] = graphClient.graphRequest.mock.calls[0];
    expect(path).toBe('/drives/drive123/items/item456/content');
    expect(options.headers['Content-Type']).toBe('application/octet-stream');
    expect(Buffer.isBuffer(options.body)).toBe(true);
    expect(options.body.toString('utf-8')).toBe('Hello, world!');
  });
});
