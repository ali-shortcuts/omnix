# Workspace and provider update

## Diagnosed regression

Excel crash logs reported an unhandled cross-thread WPF exception in SendMessage/SetBusy after provider completion. VSTO does not guarantee SynchronizationContext.Current. Async UI operations now start through the owning dispatcher with a scoped DispatcherSynchronizationContext. No WPF Application singleton is required. Chat turns are persisted before starting the request.

## Provider setup

Custom is first; local providers are last and their status panel is shown only when selected. Custom and Agent Router expose name, endpoint protocol, Base URL, API key, editable model/catalog discovery, and a synthetic connection test. A bare host URL gets /v1; explicit proxy prefixes are retained. Anthropic uses /messages, x-api-key, anthropic-version, system/max_tokens and text/image content blocks. OpenAI uses /chat/completions and bearer authentication. No silent protocol switching or credential forwarding to another host is attempted.

Test Connection sends only fixed synthetic text. It does not send document context, history or images and does not grant later document consent. Local Only still prevents a cloud test. Normal chat retains consent, once per provider/endpoint/privacy setting per request, including tool rounds.

## Official sources checked

- Anthropic SDK message/auth contract: https://github.com/anthropics/anthropic-sdk-python
- OmniRoute documents GET /v1/models with bearer authentication: https://github.com/diegosouzapw/OmniRoute
- SambaNova Base URL: https://docs.sambanova.ai/docs/en/get-started/api-keys-urls
- NVIDIA Base URL: https://docs.api.nvidia.com/nim/reference/llm-apis

Agent Router's public site did not expose usable API documentation during this check. Its editable preset uses the user-supplied origin, not a verified guarantee of service compatibility. Select the protocol specified by the account provider and run Test Connection. NVIDIA/SambaNova account eligibility, credits and quotas remain provider-dependent; they are not advertised as unlimited free services. Catalog discovery is best-effort; manual IDs remain supported.

## Verification and limits

Windows CI compiles production code and runs WPF in a fresh process without Application.Current, null-context async success/error continuations, real loopback HTTP protocol fixtures, theme/dropdown checks, and synthetic diagnostics/privacy tests. These tests do not launch licensed Excel/Word/PowerPoint and cannot certify every Office/Windows version or a customer's API key. No claim of full-document vision or a complete relational database engine is made. See OFFICE-CAPABILITIES.md for supported tools and coverage.
