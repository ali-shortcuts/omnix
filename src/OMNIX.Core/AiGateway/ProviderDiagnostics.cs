using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.Errors;
using OMNIX.Core.Storage;

namespace OMNIX.Core.AiGateway
{
    public enum ConnectionDiagnosticState
    {
        Connected,
        ReachableCatalogUnavailable,
        AuthenticationFailed,
        RateLimited,
        TimedOut,
        NetworkFailed,
        ProviderRejected
    }

    public sealed class ConnectionDiagnosticResult
    {
        public ConnectionDiagnosticState State { get; set; }
        public bool EndpointReachable { get; set; }
        public bool AuthenticationProven { get; set; }
        public bool ModelCatalogAvailable { get; set; }
        public int ModelCount { get; set; }
        public ErrorCode ErrorCode { get; set; }
        public long LatencyMs { get; set; }
        public string Summary { get; set; }
    }

    public enum ModelVerificationState
    {
        Working,
        AccessDenied,
        NotFoundOrUnavailable,
        RateLimited,
        TimedOut,
        NetworkFailed,
        Incompatible,
        PrivacyBlocked,
        Failed
    }

    public sealed class ModelVerificationResult
    {
        public string ModelId { get; set; }
        public ModelVerificationState State { get; set; }
        public ErrorCode ErrorCode { get; set; }
        public long LatencyMs { get; set; }
        public string Summary { get; set; }

        public bool Working { get { return State == ModelVerificationState.Working; } }
    }

    /// <summary>
    /// Provider diagnostics are intentionally split into three different concepts:
    /// connection/authentication, selected-model inference, and bounded catalog verification.
    /// Synthetic diagnostics never include Office document content, chat history or screenshots.
    /// </summary>
    public static class ProviderDiagnostics
    {
        private const int MaxModelsToVerify = 100;
        private static readonly TimeSpan PerModelVerificationTimeout = TimeSpan.FromSeconds(15);

        public static async Task TestModelAsync(IProviderAdapter adapter, ProviderCredentials credentials,
            PrivacyGate privacy, CancellationToken ct)
        {
            if (adapter == null) throw new ArgumentNullException("adapter");
            if (privacy == null) throw new ArgumentNullException("privacy");
            ct.ThrowIfCancellationRequested();
            adapter.Configure(credentials);
            await privacy.EnsureAllowedAsync(adapter).ConfigureAwait(true);
            ct.ThrowIfCancellationRequested();

            await SendSyntheticAsync(adapter, ct).ConfigureAwait(false);
        }

        /// <summary>
        /// Connection test only. It never sends a chat/inference request and therefore never
        /// depends on the currently selected model. For known providers, an authenticated live
        /// model-catalog request is the strongest model-independent connection/auth test available.
        /// A 404 is reported separately because some compatible endpoints omit /models.
        /// </summary>
        public static async Task<ConnectionDiagnosticResult> TestConnectionOnlyAsync(
            IProviderAdapter adapter,
            ProviderCredentials credentials,
            CancellationToken ct)
        {
            if (adapter == null) throw new ArgumentNullException("adapter");
            ct.ThrowIfCancellationRequested();

            var sw = Stopwatch.StartNew();
            try
            {
                var modelIndependent = CloneCredentials(credentials, "");
                adapter.Configure(modelIndependent);
                var models = await adapter.ListModelsAsync(ct).ConfigureAwait(false);
                sw.Stop();
                int count = models != null ? models.Count : 0;
                return new ConnectionDiagnosticResult
                {
                    State = ConnectionDiagnosticState.Connected,
                    EndpointReachable = true,
                    AuthenticationProven = true,
                    ModelCatalogAvailable = true,
                    ModelCount = count,
                    ErrorCode = ErrorCode.None,
                    LatencyMs = sw.ElapsedMilliseconds,
                    Summary = "Connection OK. Endpoint and authentication worked; model catalog returned " +
                              count + " model" + (count == 1 ? "" : "s") + ". No model was tested."
                };
            }
            catch (OperationCanceledException) { throw; }
            catch (OmnixException ex)
            {
                sw.Stop();
                string technical = ex.TechnicalDetails ?? "";
                bool httpReached = technical.IndexOf("HTTP=", StringComparison.OrdinalIgnoreCase) >= 0;

                if (ex.Code == ErrorCode.AUTH_ERROR)
                    return Result(ConnectionDiagnosticState.AuthenticationFailed, httpReached, false, false,
                        ex.Code, sw.ElapsedMilliseconds,
                        "Endpoint reached, but authentication/API-key access was rejected. No model was tested.");

                if (ex.Code == ErrorCode.MODEL_ERROR &&
                    technical.IndexOf("HTTP=404", StringComparison.OrdinalIgnoreCase) >= 0)
                    return Result(ConnectionDiagnosticState.ReachableCatalogUnavailable, true, false, false,
                        ex.Code, sw.ElapsedMilliseconds,
                        "Endpoint is reachable, but the configured Base URL does not expose a compatible /models catalog. Authentication/model access is not proven. You can still enter an exact model ID and use Test model.");

                if (ex.Code == ErrorCode.TIMEOUT)
                    return Result(ConnectionDiagnosticState.TimedOut, httpReached, false, false,
                        ex.Code, sw.ElapsedMilliseconds, "Connection test timed out before a model-independent catalog response was completed.");

                if (ex.Code == ErrorCode.NETWORK_ERROR)
                    return Result(ConnectionDiagnosticState.NetworkFailed, false, false, false,
                        ex.Code, sw.ElapsedMilliseconds, "Could not reach the provider endpoint. Check Base URL, DNS/TLS, proxy/firewall and network access.");

                if (ex.Code == ErrorCode.PROVIDER_ERROR &&
                    technical.IndexOf("rate_limit_or_quota", StringComparison.OrdinalIgnoreCase) >= 0)
                    return Result(ConnectionDiagnosticState.RateLimited, true, true, false,
                        ex.Code, sw.ElapsedMilliseconds, "Provider endpoint responded but rate/quota limits prevented the catalog check. No model was tested.");

                return Result(ConnectionDiagnosticState.ProviderRejected, httpReached, false, false,
                    ex.Code, sw.ElapsedMilliseconds, "Provider endpoint rejected the model-independent connection check. No model was tested.");
            }
        }

        // Explicit Test model click authorizes ONLY this fixed, document-free request.
        // Never reuse this path for chat, history, screenshots, or Office context.
        public static async Task TestSyntheticModelAsync(IProviderAdapter adapter, ProviderCredentials credentials, CancellationToken ct)
        {
            if (adapter == null) throw new ArgumentNullException("adapter");
            ct.ThrowIfCancellationRequested();
            adapter.Configure(credentials);
            if (adapter.Info.Kind == ProviderKind.Cloud && Settings.SettingsManager.Instance.Settings.Privacy == Settings.PrivacyMode.LocalOnly)
                throw OmnixException.PrivacyBlocked("Local Only is enabled. Select a local provider or change Privacy.");
            await SendSyntheticAsync(adapter, ct).ConfigureAwait(false);
        }

        public static async Task<ModelVerificationResult> TestSelectedModelAsync(
            string providerId,
            ProviderCredentials credentials,
            CancellationToken ct)
        {
            string model = credentials != null ? credentials.Model : null;
            if (string.IsNullOrWhiteSpace(model))
                throw OmnixException.Model("Select or enter an exact model ID before using Test model.");

            var registry = new ProviderRegistry();
            var adapter = registry.Get(providerId);
            if (adapter == null) throw OmnixException.Provider("Unknown provider: " + providerId);

            var sw = Stopwatch.StartNew();
            try
            {
                await TestSyntheticModelAsync(adapter, CloneCredentials(credentials, model), ct).ConfigureAwait(false);
                sw.Stop();
                return new ModelVerificationResult
                {
                    ModelId = model,
                    State = ModelVerificationState.Working,
                    ErrorCode = ErrorCode.None,
                    LatencyMs = sw.ElapsedMilliseconds,
                    Summary = "Working — text inference succeeded in " + sw.ElapsedMilliseconds + " ms."
                };
            }
            catch (OperationCanceledException) { throw; }
            catch (OmnixException ex)
            {
                sw.Stop();
                return ClassifyModelFailure(model, ex, sw.ElapsedMilliseconds);
            }
        }

        /// <summary>
        /// Sequential, bounded verification intentionally avoids a burst of concurrent billable
        /// requests. Each advertised model gets an isolated adapter instance and a 15-second cap.
        /// </summary>
        public static async Task<IReadOnlyList<ModelVerificationResult>> VerifyModelsAsync(
            string providerId,
            ProviderCredentials baseCredentials,
            IEnumerable<string> modelIds,
            Action<ModelVerificationResult, int, int> onResult,
            CancellationToken ct)
        {
            var models = (modelIds ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(MaxModelsToVerify)
                .ToList();

            var results = new List<ModelVerificationResult>();
            for (int i = 0; i < models.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                string model = models[i];

                ModelVerificationResult result;
                using (var perModel = new CancellationTokenSource(PerModelVerificationTimeout))
                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(ct, perModel.Token))
                {
                    try
                    {
                        result = await TestSelectedModelAsync(
                            providerId,
                            CloneCredentials(baseCredentials, model),
                            linked.Token).ConfigureAwait(false);
                    }
                    catch (OperationCanceledException)
                    {
                        if (ct.IsCancellationRequested) throw;
                        result = new ModelVerificationResult
                        {
                            ModelId = model,
                            State = ModelVerificationState.TimedOut,
                            ErrorCode = ErrorCode.TIMEOUT,
                            LatencyMs = (long)PerModelVerificationTimeout.TotalMilliseconds,
                            Summary = "Timed out."
                        };
                    }
                }

                results.Add(result);
                if (onResult != null)
                {
                    try { onResult(result, i + 1, models.Count); } catch { }
                }
            }
            return results;
        }

        private static async Task SendSyntheticAsync(IProviderAdapter adapter, CancellationToken ct)
        {
            // Model catalogs are optional on compatible servers and do not prove inference access.
            var response = await adapter.SendAsync(new ChatRequest
            {
                UserTurn = new ChatTurn
                {
                    Role = ChatRole.User,
                    Text = "Reply with OK.",
                    TimestampUtc = DateTime.UtcNow
                }
            }, null, ct).ConfigureAwait(false);
            ct.ThrowIfCancellationRequested();
            if (response == null || response.WasCancelled || string.IsNullOrWhiteSpace(response.Text))
                throw OmnixException.Provider("The selected model returned no text. Check its model ID and text/chat support.");
        }

        public static IReadOnlyList<string> ModelOptions(IEnumerable<string> discovered, string current)
        {
            var options = (discovered ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.Ordinal).Take(5000).ToList();
            // A manually entered model need not be advertised by the server's catalog.
            if (!string.IsNullOrWhiteSpace(current) && !options.Contains(current))
                options.Insert(0, current);
            return options;
        }

        private static ProviderCredentials CloneCredentials(ProviderCredentials source, string model)
        {
            return new ProviderCredentials
            {
                ApiKey = source != null ? source.ApiKey : null,
                BaseUrl = source != null ? source.BaseUrl : null,
                Model = model ?? ""
            };
        }

        private static ConnectionDiagnosticResult Result(
            ConnectionDiagnosticState state,
            bool reachable,
            bool auth,
            bool catalog,
            ErrorCode code,
            long latency,
            string summary)
        {
            return new ConnectionDiagnosticResult
            {
                State = state,
                EndpointReachable = reachable,
                AuthenticationProven = auth,
                ModelCatalogAvailable = catalog,
                ErrorCode = code,
                LatencyMs = latency,
                Summary = summary
            };
        }

        private static ModelVerificationResult ClassifyModelFailure(string model, OmnixException ex, long latency)
        {
            ModelVerificationState state;
            switch (ex.Code)
            {
                case ErrorCode.AUTH_ERROR:
                    state = ModelVerificationState.AccessDenied;
                    break;
                case ErrorCode.MODEL_ERROR:
                    state = ModelVerificationState.NotFoundOrUnavailable;
                    break;
                case ErrorCode.TIMEOUT:
                    state = ModelVerificationState.TimedOut;
                    break;
                case ErrorCode.NETWORK_ERROR:
                    state = ModelVerificationState.NetworkFailed;
                    break;
                case ErrorCode.PRIVACY_BLOCKED:
                    state = ModelVerificationState.PrivacyBlocked;
                    break;
                case ErrorCode.PROVIDER_ERROR:
                    string details = ex.TechnicalDetails ?? "";
                    if (details.IndexOf("rate_limit_or_quota", StringComparison.OrdinalIgnoreCase) >= 0)
                        state = ModelVerificationState.RateLimited;
                    else if (details.IndexOf("request_rejected", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             details.IndexOf("request_too_large", StringComparison.OrdinalIgnoreCase) >= 0)
                        state = ModelVerificationState.Incompatible;
                    else
                        state = ModelVerificationState.Failed;
                    break;
                default:
                    state = ModelVerificationState.Failed;
                    break;
            }

            string label;
            switch (state)
            {
                case ModelVerificationState.AccessDenied: label = "access denied"; break;
                case ModelVerificationState.NotFoundOrUnavailable: label = "model unavailable"; break;
                case ModelVerificationState.RateLimited: label = "rate/quota limited"; break;
                case ModelVerificationState.TimedOut: label = "timed out"; break;
                case ModelVerificationState.NetworkFailed: label = "network failed"; break;
                case ModelVerificationState.Incompatible: label = "not compatible with this text/chat request"; break;
                case ModelVerificationState.PrivacyBlocked: label = "blocked by Local Only privacy mode"; break;
                default: label = "failed"; break;
            }

            return new ModelVerificationResult
            {
                ModelId = model,
                State = state,
                ErrorCode = ex.Code,
                LatencyMs = latency,
                Summary = label + " — " + ex.Message
            };
        }
    }
}
