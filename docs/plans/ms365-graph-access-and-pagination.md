# MS365 Graph Access and Pagination Plan

## Current findings

- `get-sharepoint-sites-delta` currently returns Microsoft Graph `403 accessDenied` in AgentSmith.
- This is treated as a known access issue, not an MCP tool-shape issue.
- Discovery mode keeps prompt size small, but it must still expose important executor controls such as `fetchAllPages`.
- Personal OneDrive discovery is incomplete unless `list-drives` is enabled, because most Drive tools require a `drive-id`.

## Microsoft Graph access to add or verify

For SharePoint site delta support, verify the Azure app registration and tenant consent include the SharePoint/Files permissions needed for tenant-wide site discovery and delta reads:

- `Sites.Read.All`
- `Sites.ReadWrite.All` if write scenarios remain required
- `Files.Read.All` or `Files.ReadWrite.All` for OneDrive and document-library traversal

Also verify the signed-in/service account has access to the target SharePoint sites. If delegated permissions remain insufficient for `/sites/delta`, switch the MCP identity to an admin-consented app-only/service account flow for tenant-wide SharePoint discovery while preserving least-privilege tool allowlists.

## Pagination improvements

- Expose `fetchAllPages` in discovery-mode schemas for Graph `GET` tools.
- Expose `includeHeaders` and `excludeResponse` in discovery-mode schemas for all Graph tools.
- Run the service with conservative pagination bounds:
  - `MS365_MCP_MAX_PAGES=5`
  - `MS365_MCP_MAX_ITEMS=200`
- Agents should still start with narrow `$select`, small `$top`, and filters/search terms before using `fetchAllPages`.

## OneDrive and SharePoint write validation

- Enable `list-drives` so agents can discover the personal OneDrive drive id before calling root/list/search tools.
- Validate personal OneDrive with read-only calls first: `list-drives`, `get-drive-root-item`, and `list-folder-files`.
- Validate SharePoint write access on a disposable test location only, using clearly named artifacts such as `OpenClaw MS365 smoke test - delete me - YYYY-MM-DD`.
- Current enabled tools can create folders in SharePoint document libraries via `create-onedrive-folder`. Creating a new document/file needs a path-based upload/create-file endpoint or a corrected upload-session path shape.
