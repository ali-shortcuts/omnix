using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.Errors;

namespace OMNIX.Core.AiGateway
{
    public enum ProviderKind
    {
        Local,
        Cloud
    }

    public enum VisionSupport
    {
        No,
        Yes,
        DependsOnModel
    }

    /// <summary>
    /// Cost/access metadata is informational only. Unknown is deliberately zero/default so a
    /// provider can never be mislabeled as free merely because metadata was not initialized.
    /// </summary>
    public enum ProviderAccessProfile
    {
        Unknown = 0,
        LocalNoCost = 1,
        FreeTierAvailable = 2,
        FreeModelsAvailable = 3,
        AccountDependent = 4,
        CustomEndpoint = 5,
        FreeCreditsAvailable = 6
    }

    public sealed class ProviderInfo
    {
        public override string ToString() { return DisplayName ?? Id ?? ""; }
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public ProviderKind Kind { get; set; }
        public VisionSupport Vision { get; set; }
        public string DefaultModel { get; set; }
        public bool RequiresApiKey { get; set; }
        public string Notes { get; set; }
        public ProviderAccessProfile AccessProfile { get; set; }
        public string AccessNotes { get; set; }
        public string OfficialWebsiteUrl { get; set; }
        public string DocumentationUrl { get; set; }
        public string ApiKeyUrl { get; set; }

        /// <summary>
        /// UTC calendar date when OMNIX maintainers last checked AccessProfile/AccessNotes against
        /// an official provider page. This is deliberately visible in Settings because cloud
        /// pricing/free-tier rules can change without an OMNIX binary update.
        /// </summary>
        public string AccessVerifiedUtc { get; set; }

        /// <summary>Official source used for the current access/free-tier classification.</summary>
        public string AccessVerificationUrl { get; set; }
    }

    public interface IProviderAdapter
    {
        ProviderInfo Info { get; }
        void Configure(ProviderCredentials credentials);
        Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct);
        Task<bool> TestConnectionAsync(CancellationToken ct);
        Task<ChatResponse> SendAsync(ChatRequest request, Action<string> onDelta, CancellationToken ct);
        bool SupportsVisionNow();
    }

    /// <summary>
    /// Maps provider HTTP failures to OMNIX categories without copying the provider response body
    /// into exceptions, UI diagnostics or logs. Provider bodies are untrusted and may echo Office
    /// document text, prompts, model output, account identifiers or secrets. They are inspected only
    /// in-memory for a small number of deterministic routing signatures and then discarded.
    /// </summary>
    public static class HttpStatusMapper
    {
        private const int MaxDetectionBodyChars = 16 * 1024;

        public static OmnixException Map(int statusCode, string body, string providerName)
        {
            string detectionBody = body ?? string.Empty;
            if (detectionBody.Length > MaxDetectionBodyChars)
                detectionBody = detectionBody.Substring(0, MaxDetectionBodyChars);

            string provider = SafeProviderName(providerName);

            if (string.Equals(provider, "OpenRouter", StringComparison.OrdinalIgnoreCase) &&
                (detectionBody.IndexOf("No endpoints available matching your guardrail restrictions and data policy", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 detectionBody.IndexOf("openrouter.ai/settings/privacy", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return OmnixException.PrivacyBlocked(
                    Diagnostic(provider, statusCode, "data_policy_no_eligible_endpoint") +
                    "; settings=https://openrouter.ai/settings/privacy");
            }

            switch (statusCode)
            {
                case 401:
                    return OmnixException.Auth(Diagnostic(provider, statusCode, "authentication_failed"));
                case 403:
                    return OmnixException.Auth(Diagnostic(provider, statusCode, "access_forbidden"));
                case 404:
                    return OmnixException.Model(Diagnostic(provider, statusCode, "model_or_endpoint_not_found"));
                case 408:
                    return OmnixException.Timeout(Diagnostic(provider, statusCode, "provider_timeout"));
                case 409:
                    return OmnixException.Provider(Diagnostic(provider, statusCode, "provider_conflict"));
                case 413:
                    return OmnixException.Provider(Diagnostic(provider, statusCode, "request_too_large"));
                case 422:
                    return OmnixException.Provider(Diagnostic(provider, statusCode, "request_rejected"));
                case 429:
                    return OmnixException.Provider(Diagnostic(provider, statusCode, "rate_limit_or_quota"));
                case 500:
                case 502:
                case 503:
                case 504:
                    return OmnixException.Provider(Diagnostic(provider, statusCode, "provider_unavailable"));
                default:
                    return OmnixException.Provider(Diagnostic(provider, statusCode, "provider_http_error"));
            }
        }

        private static string Diagnostic(string provider, int statusCode, string category)
        {
            return "Provider=" + provider + "; HTTP=" + statusCode + "; category=" + category +
                   "; provider_response_body=REDACTED";
        }

        private static string SafeProviderName(string providerName)
        {
            if (string.IsNullOrWhiteSpace(providerName)) return "UnknownProvider";
            string value = providerName.Trim();
            var chars = new char[Math.Min(80, value.Length)];
            int written = 0;
            for (int i = 0; i < value.Length && written < chars.Length; i++)
            {
                char c = value[i];
                if (char.IsLetterOrDigit(c) || c == ' ' || c == '-' || c == '_' || c == '.')
                    chars[written++] = c;
            }
            if (written == 0) return "UnknownProvider";
            return new string(chars, 0, written);
        }
    }
}
