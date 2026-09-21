using System;
using System.Collections.Generic;
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
    /// Ollama adapter (local AI, port 11434). NDJSON streaming via /api/chat.
    /// Vision is enabled only when the installed model is multimodal. Request/history replay,
    /// streamed assistant output and model discovery are bounded so a local endpoint cannot grow
    /// memory without limit inside an Office host.
    /// </summary>
    public sealed class OllamaAdapter : IProviderAdapter
    {
        private const int MaxAssistantChars = 2 * 1024 * 1024;
        private const int MaxJsonBodyBytes = 8 * 1024 * 1024;
        private const int MaxImageBytes = 20 * 1024 * 1024;
        private const int MaxModels = 5000;

        private ProviderCredentials _creds;
        private static readonly string[] KnownVisionMarkers =
        {
            "llava", "bakllava", "moondream", "minicpm-v", "qwen2-vl", "qwen2.5vl",
            "llama3.2-vision", "vision", "gemma3", "mistral-small3.1", "granite3.1-moe"
        };

        public OllamaAdapter()
        {
            Info = new ProviderInfo
            {
                Id = "ollama",
                DisplayName = "Ollama (Local)",
                Kind = ProviderKind.Local,
                Vision = VisionSupport.DependsOnModel,
                DefaultModel = "",
                RequiresApiKey = false,
                Notes = "Runs on this PC (port 11434). Vision requires a multimodal local model."
            };
        }

        public ProviderInfo Info { get; private set; }

        public void Configure(ProviderCredentials credentials) { _creds = credentials ?? new ProviderCredentials(); }

        private string BaseUrl
        {
            get
            {
                return (_creds != null && !string.IsNullOrEmpty(_creds.BaseUrl))
                    ? _creds.BaseUrl.TrimEnd('/')
                    : "http://localhost:11434";
            }
        }

        private string Model
        {
            get { return _creds != null && !string.IsNullOrEmpty(_creds.Model) ? _creds.Model : FirstModelOr("llava"); }
        }

        private string _firstModel;

        private string FirstModelOr(string fallback)
        {
            return string.IsNullOrEmpty(_firstModel) ? fallback : _firstModel;
        }

        private static JObject BuildMessage(ChatTurn turn)
        {
            var msg = new JObject { { "role", turn.Role == ChatRole.Assistant ? "assistant" : "user" } };
            msg["content"] = turn.Text ?? "";
            if (turn.HasImages)
            {
                var images = new JArray();
                foreach (var img in turn.Images.Where(i => i != null && i.PngBytes != null && i.PngBytes.Length > 0))
                {
                    if (img.PngBytes.Length > MaxImageBytes)
                        throw OmnixException.Model("Image attachment exceeds the 20 MB OMNIX safety limit.");
                    images.Add(Convert.ToBase64String(img.PngBytes));
                }
                msg["images"] = images;
            }
            return msg;
        }

        private static bool HasNativeTools(ChatRequest request)
        {
            return request != null && request.Tools != null && request.Tools.Count > 0;
        }

        private static JArray BuildTools(ChatRequest request)
        {
            var tools = new JArray();
            if (!HasNativeTools(request)) return tools;
            foreach (var def in request.Tools)
            {
                if (def == null || string.IsNullOrWhiteSpace(def.Name)) continue;
                JObject schema;
                try { schema = string.IsNullOrWhiteSpace(def.ParametersJson) ? new JObject { ["type"] = "object" } : JObject.Parse(def.ParametersJson); }
                catch { schema = new JObject { ["type"] = "object" }; }
                tools.Add(new JObject
                {
                    ["type"] = "function",
                    ["function"] = new JObject
                    {
                        ["name"] = def.Name,
                        ["description"] = def.Description ?? "",
                        ["parameters"] = schema
                    }
                });
            }
            return tools;
        }

        private static List<NativeToolCall> ParseToolCalls(JObject root)
        {
            var result = new List<NativeToolCall>();
            foreach (var call in root.SelectToken("message.tool_calls") as JArray ?? new JArray())
            {
                string name = (string)call.SelectToken("function.name");
                if (string.IsNullOrWhiteSpace(name)) continue;
                var args = call.SelectToken("function.arguments");
                result.Add(new NativeToolCall
                {
                    Id = (string)call["id"] ?? Guid.NewGuid().ToString("N"),
                    Name = name,
                    ArgumentsJson = args == null ? "{}" :
                        (args.Type == JTokenType.String ? args.ToString() : args.ToString(Formatting.None))
                });
            }
            return result;
        }

        public string BuildPayload(ChatRequest request, bool stream)
        {
            request = ChatRequestBudgeter.Apply(request);

            var messages = new JArray();
            if (!string.IsNullOrEmpty(request.SystemPrompt))
                messages.Add(new JObject { { "role", "system" }, { "content", request.SystemPrompt } });
            if (request.History != null)
                foreach (var t in request.History)
                    messages.Add(BuildMessage(t));
            if (request.UserTurn != null)
                messages.Add(BuildMessage(request.UserTurn));

            var payload = new JObject
            {
                { "model", Model },
                { "messages", messages },
                { "stream", stream }
            };
            var tools = BuildTools(request);
            if (tools.Count > 0) payload["tools"] = tools;
            return payload.ToString(Formatting.None);
        }

        public async Task<ChatResponse> SendAsync(ChatRequest request, Action<string> onDelta, CancellationToken ct)
        {
            bool nativeTools = HasNativeTools(request);
            try
            {
                return await SendCoreAsync(request, nativeTools ? null : onDelta, ct).ConfigureAwait(false);
            }
            catch (OmnixException ex)
            {
                string details = ex.TechnicalDetails ?? "";
                bool rejected = nativeTools &&
                    (details.IndexOf("HTTP=400", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     details.IndexOf("HTTP=404", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     details.IndexOf("HTTP=422", StringComparison.OrdinalIgnoreCase) >= 0);
                if (!rejected) throw;

                Logging.Logger.Gateway("Native Ollama tools rejected; falling back to textual OMNIX tool protocol for this request.");
                var fallback = new ChatRequest
                {
                    SystemPrompt = request.SystemPrompt,
                    History = request.History,
                    UserTurn = request.UserTurn,
                    Tools = null
                };
                return await SendCoreAsync(fallback, onDelta, ct).ConfigureAwait(false);
            }
        }

        private async Task<ChatResponse> SendCoreAsync(ChatRequest request, Action<string> onDelta, CancellationToken ct)
        {
            try
            {
                using (var client = HttpClientFactory.Create())
                using (var req = new HttpRequestMessage(HttpMethod.Post, BaseUrl + "/api/chat"))
                {
                    bool streamResponse = onDelta != null;
                    req.Content = new StringContent(BuildPayload(request, streamResponse), Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            string err = await SseLineReader.ReadErrorBodyAsync(response, ct).ConfigureAwait(false);
                            throw HttpStatusMapper.Map((int)response.StatusCode, err, "Ollama");
                        }

                        var sb = new StringBuilder();
                        if (!streamResponse)
                        {
                            string full = await SseLineReader.ReadBodyBoundedAsync(response.Content, MaxJsonBodyBytes, ct).ConfigureAwait(false);
                            var root = JObject.Parse(full);
                            string text = (string)root.SelectToken("message.content") ?? "";
                            if (text.Length > MaxAssistantChars)
                                throw OmnixException.Provider("Ollama returned an over-sized assistant response.");
                            return new ChatResponse
                            {
                                Text = text,
                                Model = Model,
                                ToolCalls = ParseToolCalls(root)
                            };
                        }

                        using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                        {
                            foreach (string line in SseLineReader.ReadNdjsonLines(stream, ct))
                            {
                                JObject obj;
                                try { obj = JObject.Parse(line); }
                                catch { continue; }

                                string delta = (string)obj.SelectToken("message.content");
                                if (!string.IsNullOrEmpty(delta))
                                {
                                    if (sb.Length + delta.Length > MaxAssistantChars)
                                        throw OmnixException.Provider("Ollama streamed an over-sized assistant response.");
                                    sb.Append(delta);
                                    onDelta(delta);
                                }
                                if ((bool?)obj["done"] == true) break;
                            }
                        }
                        return new ChatResponse { Text = sb.ToString(), Model = Model };
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (HttpRequestException ex) { throw OmnixException.Network("Ollama: " + ex.Message); }
            catch (OmnixException) { throw; }
            catch (Exception ex) { throw OmnixException.Provider("Ollama transport failure: " + ex.Message); }
        }

        public async Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct)
        {
            try
            {
                using (var client = HttpClientFactory.Create(TimeSpan.FromSeconds(10)))
                using (var req = new HttpRequestMessage(HttpMethod.Get, BaseUrl + "/api/tags"))
                using (var response = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        string err = await SseLineReader.ReadErrorBodyAsync(response, ct).ConfigureAwait(false);
                        throw HttpStatusMapper.Map((int)response.StatusCode, err, "Ollama");
                    }

                    string json = await SseLineReader.ReadBodyBoundedAsync(response.Content, MaxJsonBodyBytes, ct).ConfigureAwait(false);
                    var root = JObject.Parse(json);
                    var list = new List<string>();
                    foreach (var m in root["models"] ?? new JArray())
                    {
                        if (list.Count >= MaxModels) break;
                        string name = (string)m["name"];
                        if (!string.IsNullOrWhiteSpace(name)) list.Add(name);
                    }
                    _firstModel = list.FirstOrDefault();
                    return list;
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (HttpRequestException ex)
            {
                throw OmnixException.Network("Ollama is not reachable at " + BaseUrl + ": " + ex.Message);
            }
            catch (OmnixException) { throw; }
            catch (Exception ex) { throw OmnixException.Provider("Ollama model discovery failure: " + ex.Message); }
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct)
        {
            try
            {
                var models = await ListModelsAsync(ct).ConfigureAwait(false);
                return models != null && models.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        public bool SupportsVisionNow()
        {
            string m = (Model ?? "").ToLowerInvariant();
            return KnownVisionMarkers.Any(marker => m.Contains(marker));
        }
    }
}
