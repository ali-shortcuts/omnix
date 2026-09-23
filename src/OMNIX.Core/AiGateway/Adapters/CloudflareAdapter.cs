using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.AiGateway.Http;
namespace OMNIX.Core.AiGateway.Adapters
{
    public sealed class CloudflareAdapter : IProviderAdapter
    {
        private OpenAiCompatibleClient _client;
        private ProviderCredentials _credentials;
        public ProviderInfo Info { get; private set; }
        public CloudflareAdapter()
        {
            Info = new ProviderInfo { Id="cloudflare", DisplayName="Cloudflare", Kind=ProviderKind.Cloud,
                RequiresApiKey=true, Vision=VisionSupport.DependsOnModel, DefaultModel="" };
        }
        public void Configure(ProviderCredentials credentials)
        {
            _credentials=credentials;
            _client=new OpenAiCompatibleClient(credentials.BaseUrl, "Cloudflare", catalogPath:
                "../models/search?format=openrouter&task=Text%20Generation&per_page=100", pagedCatalog:true);
        }
        public Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct) { return _client.ListModelsAsync(_credentials.ApiKey,ct); }
        public async Task<bool> TestConnectionAsync(CancellationToken ct) { return await ListModelsAsync(ct).ConfigureAwait(false)!=null; }
        public Task<ChatResponse> SendAsync(ChatRequest request,Action<string> delta,CancellationToken ct)
        {
            if(string.IsNullOrWhiteSpace(_credentials.Model)) throw Errors.OmnixException.Model("Select a model first.");
            return _client.SendAsync(request,_credentials.ApiKey,_credentials.Model,delta,ct);
        }
        public bool SupportsVisionNow() { return false; }
    }
}
