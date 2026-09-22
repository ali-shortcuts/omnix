# Custom provider regression audit

Baseline: 453c575f0ca703ed439b3a561acf4f9f8b17c7e2.

## Findings corrected

- A gateway returning application/json to a streaming chat request was fed to the SSE reader, discarding its text and native tool calls. Parse the declared JSON response and deliver text to the streaming UI callback once. Native tool calls use the same materialization as nonstreaming replies. Empty responses now produce an explicit provider error instead of apparent success.
- Custom adapter protocol selection depended on mutable global settings even when credentials were already captured. Snapshot ApiType with credentials, including diagnostic clones; retain the legacy settings fallback for existing callers.
- Discovery and verification conflated case-distinct model identifiers. Keep identity comparisons ordinal throughout these paths.
- The manual model textbox remained editable during a test, allowing the displayed model to differ from the one being tested. Disable it with the other request fields.
- Changing the protocol or key retained a prior vision capability result. Invalidate the cached capability when either changes.

## Verification

The Windows workspace regression harness exercises both OpenAI and Anthropic via loopback HTTP, with JSON, SSE, and JSON returned despite stream=true. It covers text and native tool calls, authentication headers, routes, and a credential protocol that disagrees with global settings. WPF checks cover manual model retention, case-distinct catalog entries, and editing while diagnostics run. Existing CI suites cover privacy, request budgets, tool execution contracts and packaging.

These fixtures do not prove a private API key has credit, access to a model, or the provider supports the selected protocol. No claim of live Excel/Word/PowerPoint testing is made by hosted CI. Some compatible endpoints expose no model catalog; a manually entered model remains supported. Responses-API-only endpoints are not Chat Completions endpoints and are outside the two protocols currently supported.
