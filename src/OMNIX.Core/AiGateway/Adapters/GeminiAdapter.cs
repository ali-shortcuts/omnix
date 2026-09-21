using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OMNIX.Core.AiGateway.Http;
using OMNIX.Core.Errors;
using OMNIX.Core.Storage;

namespace OMNIX.Core.AiGateway.Adapters
{
    /// <summary>
    /// Gemini adapter — multimodal Gemini models (image + text).
    /// Endpoint: POST /v1beta/models/{model}:streamGenerateContent?alt=sse.
    /// API key travels in x-goog-api-key (never a logged query string). Request history plus
    /// responses/model catalogs are bounded before full materialization to protect Office hosts.
    /// </summary>
    public sealed class GeminiAdapter : IProviderAdapter
    {
        private const string Base = "https://generativelanguage.googleapis.com/v1beta";
        private const int MaxAssistantChars = 2 * 1024 * 1024;
        private const int MaxJsonBodyBytes = 8 * 1024 * 1024;
        private const int MaxImageBytes = 20 * 1024 * 1024;
        private const int MaxModels = 5000;

        private ProviderCredentials _creds;

        public ProviderInfo Info { get; private set; }

        public GeminiAdapter()
        {
            Info = new ProviderInfo
            {
                Id = "gemini",
                DisplayName = "Google Gemini",
                Kind = ProviderKind.Cloud,
                Vision = VisionSupport.Yes,
                DefaultModel = "",
                RequiresApiKey = true,
                Notes = "Multimodal Gemini provider. Current access/price metadata is maintained in ProviderRegistry."
            };
        }

        public void Configure(ProviderCredentials credentials) { _creds = credentials ?? new ProviderCredentials(); }

        private string Model
        {
            get { return _creds == null || string.IsNullOrEmpty(_creds.Model) ? Info.DefaultModel : _creds.Model; }
        }

        private static JObject BuildPart(ChatTurn turn)
        {
            var parts = new JArray();
            if (!string.IsNullOrEmpty(turn.Text))
                parts.Add(new JObject { { "text", turn.Text } });
            if (turn.HasImages)
            {
                foreach (var img in turn.Images.Where(i => i != null && i.PngBytes != null && i.PngBytes.Length > 0))
                {
                    if (img.PngBytes.Length > MaxImageBytes)
                        throw OmnixException.Model("Image attachment exceeds the 20 MB OMNIX safety limit.");
                    parts.Add(new JObject
                    {
                        { "inline_data", new JObject { { "mime_type", "image/png" },
                            { "data", Convert.ToBase64String(img.PngBytes) } } }
                    });
                }
            }
            return new JObject { { "parts", parts } };
        }

        public string BuildPayload(ChatRequest request, bool stream)
        {
            request = ChatRequestBudgeter.Apply(request);

            var contents = new JArray();
            if (request.History != null)
                foreach (var t in request.History)
                {
                    var part = BuildPart(t);
                    part["role"] = t.Role == ChatRole.Assistant ? "model" : "user";
                    contents.Add(part);
                }
            if (request.UserTurn != null)
            {
                var part = BuildPart(request.UserTurn);
                part["role"] = "user";
                contents.Add(part);
            }

            var payload = new JObject { { "contents", contents } };
            if (!string.IsNullOrEmpty(request.SystemPrompt))
                payload["systemInstruction"] = new JObject
                {
                    { "parts", new JArray { new JObject { { "text", request.SystemPrompt } } } }
                };
            return payload.ToString(Formatting.None);
        }

        public async Task<ChatResponse> SendAsync(ChatRequest request, Action<string> onDelta, CancellationToken ct)
        {
            if (_creds == null || string.IsNullOrEmpty(_creds.ApiKey))
                throw OmnixException.Auth("No Gemini API key configured.");

            if (string.IsNullOrWhiteSpace(Model))
                throw OmnixException.Model("Select a Gemini model from Load Models or enter a model ID before sending.");

            string url = Base + "/models/" + Uri.EscapeDataString(Model) +
                         (onDelta != null ? ":streamGenerateContent?alt=sse" : ":generateContent");

            try
            {
                using (var client = HttpClientFactory.Create())
                using (var req = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    req.Headers.TryAddWithoutValidation("x-goog-api-key", _creds.ApiKey);
                    req.Content = new StringContent(BuildPayload(request, onDelta != null), Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            string err = await SseLineReader.ReadErrorBodyAsync(response, ct).ConfigureAwait(false);
                            throw HttpStatusMapper.Map((int)response.StatusCode, err, "Gemini");
                        }

                        var sb = new StringBuilder();
                        if (onDelta == null)
                        {
                            string full = await ReadBodyBoundedAsync(response.Content, MaxJsonBodyBytes, ct).ConfigureAwait(false);
                            var root = JObject.Parse(full);
                            foreach (var part in root.SelectTokens("candidates[0].content.parts[*]"))
                            {
                                string text = (string)part["text"];
                                if (string.IsNullOrEmpty(text)) continue;
                                if (sb.Length + text.Length > MaxAssistantChars)
                                    throw OmnixException.Provider("Gemini returned an over-sized assistant response.");
                                sb.Append(text);
                            }
                        }
                        else
                        {
                            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                            {
                                foreach (string data in SseLineReader.ReadDataLines(stream, ct))
                                {
                                    JObject chunk;
                                    try { chunk = JObject.Parse(data); }
                                    catch { continue; }
                                    foreach (var part in chunk.SelectTokens("candidates[0].content.parts[*]"))
                                    {
                                        string delta = (string)part["text"];
                                        if (!string.IsNullOrEmpty(delta))
                                        {
                                            if (sb.Length + delta.Length > MaxAssistantChars)
                                                throw OmnixException.Provider("Gemini streamed an over-sized assistant response.");
                                            sb.Append(delta);
                                            onDelta(delta);
                                        }
                                    }
                                }
                            }
                        }
                        return new ChatResponse { Text = sb.ToString(), Model = Model };
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (HttpRequestException ex) { throw OmnixException.Network("Gemini: " + ex.Message); }
            catch (OmnixException) { throw; }
            catch (Exception ex) { throw OmnixException.Provider("Gemini transport failure: " + ex.Message); }
        }

        public async Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct)
        {
            if (_creds == null || string.IsNullOrEmpty(_creds.ApiKey))
                throw OmnixException.Auth("No Gemini API key configured.");

            try
            {
                var list = new List<string>();
                string pageToken = null;
                int pages = 0;

                using (var client = HttpClientFactory.Create(TimeSpan.FromSeconds(20)))
                {
                    do
                    {
                        ct.ThrowIfCancellationRequested();
                        string url = Base + "/models?pageSize=1000" +
                                     (string.IsNullOrEmpty(pageToken) ? "" : "&pageToken=" + Uri.EscapeDataString(pageToken));
                        using (var req = new HttpRequestMessage(HttpMethod.Get, url))
                        {
                            req.Headers.TryAddWithoutValidation("x-goog-api-key", _creds.ApiKey);
                            using (var response = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                            {
                                if (!response.IsSuccessStatusCode)
                                {
                                    string err = await SseLineReader.ReadErrorBodyAsync(response, ct).ConfigureAwait(false);
                                    throw HttpStatusMapper.Map((int)response.StatusCode, err, "Gemini");
                                }

                                string json = await ReadBodyBoundedAsync(response.Content, MaxJsonBodyBytes, ct).ConfigureAwait(false);
                                var root = JObject.Parse(json);
                                foreach (var m in root["models"] ?? new JArray())
                                {
                                    if (list.Count >= MaxModels) break;
                                    string name = (string)m["name"] ?? "";
                                    if (name.StartsWith("models/", StringComparison.Ordinal)) name = name.Substring(7);
                                    var methods = m["supportedGenerationMethods"] as JArray;
                                    if (methods != null && !methods.Any(t => (string)t == "generateContent")) continue;
                                    if (!name.StartsWith("gemini-", StringComparison.OrdinalIgnoreCase)) continue;
                                    if (!string.IsNullOrWhiteSpace(name) && !list.Contains(name, StringComparer.OrdinalIgnoreCase))
                                        list.Add(name);
                                }
                                pageToken = list.Count >= MaxModels ? null : (string)root["nextPageToken"];
                            }
                        }
                        pages++;
                    }
                    while (!string.IsNullOrEmpty(pageToken) && pages < 5);
                }
                return list;
            }
            catch (OperationCanceledException) { throw; }
            catch (HttpRequestException ex) { throw OmnixException.Network("Gemini models: " + ex.Message); }
            catch (OmnixException) { throw; }
            catch (Exception ex) { throw OmnixException.Provider("Gemini model discovery failure: " + ex.Message); }
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct)
        {
            try
            {
                var models = await ListModelsAsync(ct).ConfigureAwait(false);
                return models != null;
            }
            catch
            {
                return false;
            }
        }

        public bool SupportsVisionNow() { return true; }

        private static async Task<string> ReadBodyBoundedAsync(HttpContent content, int maxBytes, CancellationToken ct)
        {
            if (content == null) return string.Empty;
            using (var stream = await content.ReadAsStreamAsync().ConfigureAwait(false))
            using (var ms = new MemoryStream())
            {
                var buffer = new byte[8192];
                while (true)
                {
                    ct.ThrowIfCancellationRequested();
                    int remaining = maxBytes + 1 - (int)ms.Length;
                    if (remaining <= 0) throw new InvalidDataException("Gemini response exceeded OMNIX safety limit.");
                    int read = await stream.ReadAsync(buffer, 0, Math.Min(buffer.Length, remaining), ct).ConfigureAwait(false);
                    if (read <= 0) break;
                    ms.Write(buffer, 0, read);
                    if (ms.Length > maxBytes) throw new InvalidDataException("Gemini response exceeded OMNIX safety limit.");
                }
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
    }
}
