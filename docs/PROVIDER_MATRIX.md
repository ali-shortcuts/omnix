# OMNIX Provider Matrix

Last reviewed: 2026-09-08

This document is a maintenance record, not a permanent pricing promise. Cloud providers can change models, quotas, regions, free tiers, trial credits and data policies at any time. OMNIX therefore loads model lists dynamically where the provider supports it and surfaces rate-limit/provider errors rather than silently falling back to a paid model.

| Provider | OMNIX ID | Type | API key | Free / local status | Default | Vision in OMNIX | Official setup/docs |
|---|---|---|---|---|---|---|---|
| Ollama | `ollama` | Local | No | Local; no cloud token billing | first detected local model | Model-dependent | https://docs.ollama.com/ |
| LM Studio | `lmstudio` | Local | No | Local; no cloud token billing | loaded local model | Model-dependent | https://lmstudio.ai/docs |
| Google Gemini | `gemini` | Cloud | Yes | Free tier currently available for supported models; limits/data-use/region apply | `gemini-3.8-flash` | Yes | https://ai.google.dev/gemini-api/docs / https://aistudio.google.com/apikey |
| Groq | `groq` | Cloud | Yes | Free Plan currently published with model-specific limits | `openai/gpt-oss-120b` | Model-dependent | https://console.groq.com/docs/quickstart / https://console.groq.com/keys |
| OpenRouter | `openrouter` | Cloud | Yes | `openrouter/free` + live zero-price/`:free` variants | `openrouter/free` | Model-dependent; free router supports capability routing | https://openrouter.ai/docs / https://openrouter.ai/settings/keys |
| Mistral AI | `mistral` | Cloud | Yes | Studio Free mode currently available with limited included usage/rate limits | `mistral-small-latest` | Model-dependent | https://docs.mistral.ai/getting-started/quickstarts/studio/activate-and-generate-api-key / https://console.mistral.ai/ |
| Hugging Face Inference Providers | `huggingface` | Cloud router | Yes | Small monthly free-user credits; live catalog can expose temporary free provider routes | `openai/gpt-oss-120b:fastest` | Model-dependent; detected from live model architecture metadata | https://huggingface.co/docs/inference-providers/index / https://huggingface.co/settings/tokens |
| Cerebras | `cerebras` | Cloud | Yes | Free trial credits currently advertised; continued usage is account/plan dependent | `gpt-oss-120b` | Text-only in this OMNIX build | https://www.cerebras.ai/inference / https://cloud.cerebras.ai/ |
| Custom OpenAI-compatible | `custom` | User endpoint | Optional | Defined entirely by endpoint owner; localhost is treated as local | user configured | Probe-dependent | User configured; OMNIX does not invent a setup URL |

## Free-first behavior

- Local AI is preferred when `Prefer Local AI when available` is enabled and a compatible local model is reachable.
- `Local Only` privacy mode blocks cloud sends, while a selected loopback Custom endpoint can be used as local AI after validation.
- OpenRouter model discovery places `openrouter/free` first, then currently detected zero-price / `:free` model ids, then other models.
- Hugging Face model discovery prioritizes models that the live `/v1/models` catalog reports with at least one live provider route marked `is_free=true`; its account-level monthly free credits remain separate from those promotions.
- OMNIX does not automatically convert a provider failure into a paid request.
- Cloud free-tier/free-credit/free-trial labels are informational. A `429` remains a quota/rate-limit error; a provider-policy failure remains a provider/privacy error.

## Secrets and privacy

- API keys are stored with Windows DPAPI (`CurrentUser`) and are not logged.
- Cloud sends are subject to OMNIX Privacy Mode before the provider call.
- Official setup links exposed by Settings are HTTPS and restricted to a hard-coded provider-owned host allowlist.
- Custom provider URLs are user configuration and are not opened as trusted setup links by OMNIX.
- A local Custom endpoint is treated as local only because the URL is loopback; OMNIX cannot guarantee that third-party local server software will not itself forward data elsewhere.

## Runtime release gate

Provider code compiling is not enough for release. Before OMNIX v3 is considered complete, test at least:

1. API-key save/reload without plaintext leakage.
2. Load Models against each configured cloud provider.
3. Text streaming round-trip for Gemini, Groq, OpenRouter, Mistral, Hugging Face and Cerebras.
4. Correct `401/403`, `429`, timeout and provider-policy error classification.
5. Vision request on a Vision-capable Gemini/OpenRouter/Hugging Face/Mistral model where the live catalog/model actually supports it.
6. Local-only request with the Internet disconnected and a local model running.
7. OpenRouter `openrouter/free` request and at least one live `:free` model when available.
8. Hugging Face model discovery and one routed request using remaining account credits; verify that no paid fallback is falsely claimed as free.
9. Custom provider once without an API key (local/no-auth endpoint) and once with an API key (authenticated OpenAI-compatible endpoint).
10. Verify every Settings `Get API Key` / docs button opens only the intended provider-owned HTTPS host.

Do not mark an unexecuted provider test PASS.


## Connection vs model diagnostics

OMNIX does not treat a model catalog as proof of usable inference.

- **Test connection** checks endpoint/authentication/catalog without using a model.
- **Detect models** lists the live advertised catalog.
- **Test model** verifies the selected model with a tiny synthetic text request.
- **Verify models** can test up to 100 detected IDs sequentially with bounded per-model timeouts and
  exposes which IDs actually work for the current provider/API configuration.

A provider can therefore be connected while a specific model is unavailable, denied, rate-limited,
or incompatible. Those states are intentionally reported separately.

