using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.Errors;
using OMNIX.Core.Storage;

namespace OMNIX.Core.AiGateway
{
    /// <summary>Tests the selected model with synthetic text, without reading Office content.</summary>
    public static class ProviderDiagnostics
    {
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

        // Explicit Test Connection click authorizes ONLY this fixed, document-free request.
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
                throw OmnixException.Provider("The selected model returned no text. Check its model ID and chat-completions support.");
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
    }
}
