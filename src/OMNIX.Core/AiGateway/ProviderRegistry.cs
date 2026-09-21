using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.AiGateway.Adapters;

namespace OMNIX.Core.AiGateway
{
    /// <summary>
    /// Provider registry. Provider-owned setup URLs, access/cost hints and maintained default
    /// model ids are centralized here. Cloud free-tier metadata is informational: OMNIX still
    /// loads live model lists and surfaces provider quota/rate-limit errors instead of promising
    /// that a cloud provider will remain free forever. Every cloud access classification carries
    /// an official verification source and date so stale claims remain visible to the user.
    /// </summary>
    public sealed class ProviderRegistry
    {
        private const string AccessVerifiedDate = "2026-09-09";

        private readonly List<IProviderAdapter> _providers;
        private readonly Dictionary<string, bool> _localAvailability;
        private readonly Dictionary<string, string> _localModelHints = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public ProviderRegistry()
        {
            _providers = new List<IProviderAdapter>
            {
                new CustomOpenAiCompatibleAdapter(),
                new CustomOpenAiCompatibleAdapter("agentrouter"),
                new CompatiblePresetAdapter("sambanova", "SambaNova", "https://api.sambanova.ai/v1", "https://docs.sambanova.ai/docs/en/get-started/api-keys-urls", "https://cloud.sambanova.ai/"),
                new CompatiblePresetAdapter("nvidia", "NVIDIA", "https://integrate.api.nvidia.com/v1", "https://docs.api.nvidia.com/nim/reference/llm-apis", "https://build.nvidia.com/"),
                new GeminiAdapter(),
                new GroqAdapter(),
                new OpenRouterAdapter(),
                new MistralAdapter(),
                new HuggingFaceAdapter(),
                new CerebrasAdapter(),
                new OllamaAdapter(),
                new LmStudioAdapter()
            };
            _localAvailability = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            ApplyOfficialMetadata();
            var names = new Dictionary<string, string> {
                { "custom", "Custom Provider" }, { "agentrouter", "Agent Router" }, { "sambanova", "SambaNova" }, { "nvidia", "NVIDIA" }, { "gemini", "Gemini" }, { "groq", "Groq" },
                { "openrouter", "OpenRouter" }, { "mistral", "Mistral" },
                { "huggingface", "Hugging Face" }, { "cerebras", "Cerebras" },
                { "ollama", "Ollama" }, { "lmstudio", "LM Studio" }
            };
            foreach (var provider in _providers) provider.Info.DisplayName = names[provider.Info.Id];
        }

        private void ApplyOfficialMetadata()
        {
            SetMetadata("gemini",
                "https://ai.google.dev/gemini-api/docs",
                "https://ai.google.dev/gemini-api/docs/get-started",
                "https://aistudio.google.com/apikey",
                "",
                ProviderAccessProfile.FreeTierAvailable,
                "Gemini model access, regional availability, quotas and billing depend on your Google project. Load the live model catalog and select an available model.",
                "https://ai.google.dev/gemini-api/docs/pricing",
                AccessVerifiedDate);

            SetMetadata("groq",
                "https://groq.com/",
                "https://console.groq.com/docs/quickstart",
                "https://console.groq.com/keys",
                "openai/gpt-oss-120b",
                ProviderAccessProfile.FreeTierAvailable,
                "Groq currently publishes Free Plan rate limits for openai/gpt-oss-120b and other supported models. Quotas are account/model specific.",
                "https://console.groq.com/docs/rate-limits",
                AccessVerifiedDate);

            SetMetadata("openrouter",
                "https://openrouter.ai/",
                "https://openrouter.ai/docs",
                "https://openrouter.ai/settings/keys",
                "openrouter/free",
                ProviderAccessProfile.FreeModelsAvailable,
                "OpenRouter currently exposes openrouter/free plus individual :free variants. Free capacity, model inventory and provider privacy compatibility can change.",
                "https://openrouter.ai/collections/free-models",
                AccessVerifiedDate);

            SetMetadata("mistral",
                "https://mistral.ai/",
                "https://docs.mistral.ai/getting-started/quickstarts/studio/activate-and-generate-api-key",
                "https://console.mistral.ai/",
                "mistral-small-latest",
                ProviderAccessProfile.FreeTierAvailable,
                "Mistral currently documents Free mode with API-key access, included usage and rate limits; no credit card is required for Free mode. Availability and included usage can change.",
                "https://docs.mistral.ai/admin/billing-usage/subscriptions",
                AccessVerifiedDate);

            SetMetadata("huggingface",
                "https://huggingface.co/",
                "https://huggingface.co/docs/inference-providers/index",
                "https://huggingface.co/settings/tokens",
                "openai/gpt-oss-120b:fastest",
                ProviderAccessProfile.FreeCreditsAvailable,
                "Hugging Face currently gives free users a small monthly Inference Providers credit allowance; the amount is explicitly subject to change. Live routing/model availability is still checked at runtime.",
                "https://huggingface.co/docs/inference-providers/pricing",
                AccessVerifiedDate);

            SetMetadata("cerebras",
                "https://www.cerebras.ai/",
                "https://inference-docs.cerebras.ai/",
                "https://cloud.cerebras.ai/",
                "gpt-oss-120b",
                ProviderAccessProfile.FreeCreditsAvailable,
                "Cerebras currently advertises a free trial credit allowance for new accounts and separately publishes Free-tier rate limits. Continued usage beyond trial/free capacity is account/plan dependent.",
                "https://www.cerebras.ai/pricing",
                AccessVerifiedDate);

            SetMetadata("ollama",
                "https://ollama.com/",
                "https://docs.ollama.com/",
                null,
                null,
                ProviderAccessProfile.LocalNoCost,
                "Runs locally on this PC. OMNIX does not charge or meter local inference; compute/storage use comes from the user's own machine.",
                "https://docs.ollama.com/",
                AccessVerifiedDate);

            SetMetadata("lmstudio",
                "https://lmstudio.ai/",
                "https://lmstudio.ai/docs",
                null,
                null,
                ProviderAccessProfile.LocalNoCost,
                "Runs locally on this PC through the LM Studio local server. OMNIX does not meter local inference.",
                "https://lmstudio.ai/docs",
                AccessVerifiedDate);

            var custom = Get("custom");
            if (custom != null)
            {
                custom.Info.AccessProfile = ProviderAccessProfile.CustomEndpoint;
                custom.Info.AccessNotes = "Cost, privacy, authentication and limits are defined entirely by the user-configured endpoint.";
                custom.Info.AccessVerifiedUtc = null;
                custom.Info.AccessVerificationUrl = null;
            }

            var gemini = Get("gemini");
            if (gemini != null)
                gemini.Info.Notes = "Vision-capable Gemini provider. Models are loaded dynamically; free-tier eligibility still depends on the Google account/region and current provider policy.";

            var groq = Get("groq");
            if (groq != null)
                groq.Info.Notes = "Fast GroqCloud inference. Available models and account limits are loaded/validated at runtime where possible.";

            var openRouter = Get("openrouter");
            if (openRouter != null)
                openRouter.Info.Notes = "Multi-provider router. openrouter/free and :free variants are prioritized in the model list; account privacy policy can restrict routing.";
        }

        private void SetMetadata(
            string id,
            string website,
            string docs,
            string apiKey,
            string defaultModel,
            ProviderAccessProfile accessProfile,
            string accessNotes,
            string verificationUrl,
            string verifiedUtc)
        {
            var provider = Get(id);
            if (provider == null) return;
            provider.Info.OfficialWebsiteUrl = website;
            provider.Info.DocumentationUrl = docs;
            provider.Info.ApiKeyUrl = apiKey;
            provider.Info.AccessProfile = accessProfile;
            provider.Info.AccessNotes = accessNotes;
            provider.Info.AccessVerificationUrl = verificationUrl;
            provider.Info.AccessVerifiedUtc = verifiedUtc;
            if (!string.IsNullOrWhiteSpace(defaultModel)) provider.Info.DefaultModel = defaultModel;
        }

        public IReadOnlyList<IProviderAdapter> All { get { return _providers; } }

        public IProviderAdapter Get(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return _providers.FirstOrDefault(p => string.Equals(p.Info.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        public void SetLocalAvailability(string id, bool available)
        {
            lock (_localAvailability)
            {
                _localAvailability[id] = available;
            }
        }

        public bool IsLocalAvailable(string id)
        {
            lock (_localAvailability)
            {
                bool v;
                return _localAvailability.TryGetValue(id, out v) && v;
            }
        }

        public IProviderAdapter GetFirstAvailableLocal()
        {
            foreach (var p in _providers.Where(x => x.Info.Kind == ProviderKind.Local))
            {
                if (IsLocalAvailable(p.Info.Id)) return p;
            }
            return null;
        }

        public string GetLocalModelHint(string id)
        {
            lock (_localModelHints)
            {
                string hint;
                return _localModelHints.TryGetValue(id, out hint) ? hint : null;
            }
        }

        public async Task RefreshLocalModelHintAsync(string id, CancellationToken ct)
        {
            var provider = Get(id);
            if (provider == null) return;
            var models = await provider.ListModelsAsync(ct).ConfigureAwait(false);
            lock (_localModelHints)
                _localModelHints[id] = models != null && models.Count > 0 ? models[0] : null;
        }
    }
}
