using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.AiGateway.Http;

namespace OMNIX.Core.AiGateway.Adapters
{
    public sealed class CompatiblePresetAdapter : IProviderAdapter
    {
        private readonly OpenAiCompatibleClient _client;
        private ProviderCredentials _credentials = new ProviderCredentials();
        public ProviderInfo Info { get; private set; }
        public CompatiblePresetAdapter(string id, string name, string url, string docs, string keys)
        {
            _client = new OpenAiCompatibleClient(url, name);
            Info = new ProviderInfo { Id=id, DisplayName=name, Kind=ProviderKind.Cloud,
                Vision=VisionSupport.DependsOnModel, RequiresApiKey=true, DefaultModel="",
                AccessProfile=ProviderAccessProfile.AccountDependent, DocumentationUrl=docs, ApiKeyUrl=keys };
        }
        public void Configure(ProviderCredentials credentials) { _credentials=credentials ?? new ProviderCredentials(); }
        public Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct) { return _client.ListModelsAsync(_credentials.ApiKey,ct); }
        public Task<ChatResponse> SendAsync(ChatRequest request, Action<string> onDelta, CancellationToken ct)
        {
            if(string.IsNullOrWhiteSpace(_credentials.Model)) throw Errors.OmnixException.Model("Select a model first.");
            return _client.SendAsync(request,_credentials.ApiKey,_credentials.Model,onDelta,ct);
        }
        public async Task<bool> TestConnectionAsync(CancellationToken ct) {
            await ProviderDiagnostics.TestSyntheticModelAsync(this,_credentials,ct).ConfigureAwait(false); return true;
        }
        public bool SupportsVisionNow() { return false; } // Text-tested presets; no unverified vision claim.
    }
}
