using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.AiGateway.Http;
using OMNIX.Core.Errors;
using OMNIX.Core.Settings;
using OMNIX.Core.Storage;

namespace OMNIX.Core.AiGateway.Adapters
{
    /// <summary>
    /// Custom provider: any OpenAI-compatible endpoint (Name + Base URL + optional API Key + Model).
    /// Loopback endpoints (localhost / 127.0.0.1 / ::1) are classified as Local after Configure,
    /// so Privacy Mode can use a local custom server without pretending data leaves the PC.
    /// Remote custom endpoints MUST use HTTPS because Office context and optional API credentials
    /// must never be sent over clear-text HTTP. All remote endpoints remain Cloud and pass through
    /// the normal cloud privacy gate.
    /// </summary>
    public sealed class CustomOpenAiCompatibleAdapter : IProviderAdapter
    {
        private ProviderCredentials _creds;
        private bool? _visionProbeResult;
        private string _baseUrl;
        private bool _anthropic;

        public CustomOpenAiCompatibleAdapter(string id = "custom")
        {
            Info = new ProviderInfo
            {
                Id = id,
                DisplayName = "Custom (OpenAI-compatible)",
                Kind = ProviderKind.Cloud,
                Vision = VisionSupport.DependsOnModel,
                DefaultModel = "",
                RequiresApiKey = true,
                AccessProfile = ProviderAccessProfile.CustomEndpoint,
                AccessNotes = "API key is optional. Loopback http/https endpoints are local; every non-loopback custom endpoint must use HTTPS and is treated as cloud.",
                Notes = "Configure an OpenAI-compatible Base URL, model, and optional API key."
            };
        }

        public ProviderInfo Info { get; private set; }

        public void Configure(ProviderCredentials credentials)
        {
            _creds = credentials ?? new ProviderCredentials();
            string url = _creds.BaseUrl;
            var cp = SettingsManager.Instance.Settings.EndpointConfig(Info.Id);
            if (string.IsNullOrWhiteSpace(url) && cp != null) url = cp.BaseUrl;

            _anthropic = cp != null && cp.ApiType == "Anthropic";
            Uri parsed = new Uri(NormalizeBaseUrl(url));
            _baseUrl = parsed.AbsoluteUri.TrimEnd('/');
            Info.Kind = IsLoopbackEndpoint(parsed) ? ProviderKind.Local : ProviderKind.Cloud;
            Info.AccessNotes = Info.Kind == ProviderKind.Local
                ? "Loopback custom endpoint: treated as Local AI; request data stays on this PC unless that local server forwards it elsewhere."
                : "Remote HTTPS custom endpoint: treated as Cloud; OMNIX cloud privacy rules apply. Cost, retention and limits are defined by the endpoint owner.";

            if (string.IsNullOrWhiteSpace(_creds.Model) && cp != null) _creds.Model = cp.Model;
            if (string.IsNullOrWhiteSpace(_creds.Model)) _creds.Model = Info.DefaultModel;
        }

        public Task<ChatResponse> SendAsync(ChatRequest request, Action<string> onDelta, CancellationToken ct)
        {
            return ActiveClient().SendAsync(request, _creds != null ? _creds.ApiKey : null,
                _creds != null ? _creds.Model : Info.DefaultModel, onDelta, ct);
        }

        public Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct)
        {
            return ActiveClient().ListModelsAsync(_creds != null ? _creds.ApiKey : null, ct);
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct)
        {
            try
            {
                var models = await ListModelsAsync(ct).ConfigureAwait(false);
                bool ok = models != null && models.Count > 0;
                if (ok)
                {
                    try
                    {
                        byte[] png = Convert.FromBase64String(
                            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
                        var probe = new ChatRequest
                        {
                            UserTurn = new ChatTurn
                            {
                                Role = ChatRole.User,
                                Text = "Reply with OK.",
                                Images = new List<ImageAttachment>
                                {
                                    new ImageAttachment { PngBytes = png, FileName = "probe.png" }
                                },
                                TimestampUtc = DateTime.UtcNow
                            }
                        };
                        var resp = await ActiveClient().SendAsync(probe,
                            _creds != null ? _creds.ApiKey : null,
                            _creds != null ? _creds.Model : Info.DefaultModel,
                            null, ct).ConfigureAwait(false);
                        _visionProbeResult = resp != null && resp.Text != null;
                    }
                    catch
                    {
                        _visionProbeResult = false;
                    }

                    var cp = SettingsManager.Instance.Settings.EndpointConfig(Info.Id);
                    if (cp != null) cp.SupportsVision = _visionProbeResult;
                }
                return ok;
            }
            catch
            {
                return false;
            }
        }

        public bool SupportsVisionNow()
        {
            if (_visionProbeResult.HasValue) return _visionProbeResult.Value;
            var cp = SettingsManager.Instance.Settings.EndpointConfig(Info.Id);
            return cp != null && cp.SupportsVision == true;
        }

        public static bool IsLoopbackEndpoint(string raw)
        {
            Uri uri;
            return Uri.TryCreate(raw, UriKind.Absolute, out uri) && IsLoopbackEndpoint(uri);
        }

        private static bool IsLoopbackEndpoint(Uri uri)
        {
            return uri != null && (uri.IsLoopback ||
                string.Equals(uri.Host, "localhost", StringComparison.OrdinalIgnoreCase));
        }

        private static Uri ValidateEndpoint(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw OmnixException.Provider("Custom provider Base URL is not configured. Set it in Settings.");

            Uri parsed;
            if (!Uri.TryCreate(raw.Trim(), UriKind.Absolute, out parsed) ||
                !(string.Equals(parsed.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(parsed.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
            {
                throw OmnixException.Provider("Custom provider Base URL must be an absolute http:// or https:// URL.");
            }

            if (!string.IsNullOrEmpty(parsed.UserInfo))
                throw OmnixException.Provider("Custom provider Base URL must not embed a username or password. Store the API key in OMNIX Settings instead.");

            bool loopback = IsLoopbackEndpoint(parsed);
            if (!loopback && !string.Equals(parsed.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                throw OmnixException.Provider(
                    "Remote custom providers must use HTTPS. Clear-text HTTP is allowed only for localhost/loopback endpoints.");
            }

            if (!string.IsNullOrEmpty(parsed.Fragment))
                throw OmnixException.Provider("Custom provider Base URL must not contain a URL fragment (#...).");

            if (!string.IsNullOrEmpty(parsed.Query))
                throw OmnixException.Provider("Custom provider Base URL must not contain query parameters. Store credentials in the API key field.");

            return parsed;
        }

        public static string NormalizeBaseUrl(string raw)
        {
            var parsed = ValidateEndpoint(raw);
            string value = parsed.GetLeftPart(UriPartial.Path).TrimEnd('/');
            foreach (string suffix in new[] { "/chat/completions", "/messages", "/models" })
            {
                if (value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                    return value.Substring(0, value.Length - suffix.Length);
            }
            if (parsed.AbsolutePath == "/") value += "/v1";
            return value;
        }

        private OpenAiCompatibleClient ActiveClient()
        {
            if (string.IsNullOrWhiteSpace(_baseUrl))
            {
                string raw = _creds != null ? _creds.BaseUrl : null;
                if (string.IsNullOrWhiteSpace(raw) && SettingsManager.Instance.Settings.EndpointConfig(Info.Id) != null)
                    raw = SettingsManager.Instance.Settings.EndpointConfig(Info.Id).BaseUrl;
                Uri parsed = new Uri(NormalizeBaseUrl(raw));
                _baseUrl = parsed.AbsoluteUri.TrimEnd('/');
                Info.Kind = IsLoopbackEndpoint(parsed) ? ProviderKind.Local : ProviderKind.Cloud;
            }
            return new OpenAiCompatibleClient(_baseUrl, "Custom", null, _anthropic);
        }
    }
}
