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
using OMNIX.Core.Tools;
using OMNIX.Core.Logging;

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

        private static JObject BuildNativeTool()
        {
            var names = new JArray();
            foreach (string name in ToolNames.AllWhitelisted) names.Add(name);
            return new JObject
            {
                { "type", "function" },
                { "function", new JObject
                    {
                        { "name", "omnix_tool" },
                        { "description", "Execute exactly one approved OMNIX Office tool. Use one tool call per turn." },
                        { "parameters", new JObject
                            {
                                { "type", "object" },
                                { "properties", new JObject
                                    {
                                        { "tool", new JObject { { "type", "string" }, { "enum", names } } },
                                        { "args", new JObject { { "type", "object" } } }
                                    }
                                },
                                { "required", new JArray("tool", "args") }
                            }
                        }
                    }
                }
            };
        }

        private static ProviderToolCall DecodeOllamaToolCall(JObject call)
        {
            if (call == null) return null;
            string functionName = (string)call.SelectToken("function.name") ?? "";
            JToken arguments = call.SelectToken("function.arguments");

            if (string.Equals(functionName, "omnix_tool", StringComparison.OrdinalIgnoreCase))
            {
                var root = arguments as JObject;
                if (root == null && arguments != null && arguments.Type == JTokenType.String)
                {
                    try { root = JObject.Parse((string)arguments); } catch { }
                }
                root = root ?? new JObject();
                string tool = ToolNames.Normalize((string)root["tool"] ?? "");
                JToken args = root["args"];
                return new ProviderToolCall
                {
                    Id = "",
                    Name = tool,
                    ArgumentsJson = args == null ? "{}" :
                        args.Type == JTokenType.String ? (string)args : args.ToString(Formatting.None)
                };
            }

            return new ProviderToolCall
            {
                Id = "",
                Name = ToolNames.Normalize(functionName),
                ArgumentsJson = arguments == null ? "{}" :
                    arguments.Type == JTokenType.String ? (string)arguments : arguments.ToString(Formatting.None)
            };
        }

        public string BuildPayload(ChatRequest request)
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
                { "stream", true }
            };
            if (request.UseNativeTools)
                payload["tools"] = new JArray(BuildNativeTool());
            return payload.ToString(Formatting.None);
        }

        public async Task<ChatResponse> SendAsync(ChatRequest request, Action<string> onDelta, CancellationToken ct)
        {
            try
            {
                return await SendOnceAsync(request, onDelta, ct).ConfigureAwait(false);
            }
            catch (OmnixException ex)
            {
                if (request != null && request.UseNativeTools && IsNativeToolSchemaRejection(ex))
                {
                    Logger.Gateway("Ollama rejected native tool schema; retrying this turn with OMNIX text-protocol fallback.");
                    var fallback = new ChatRequest
                    {
                        SystemPrompt = request.SystemPrompt,
                        History = request.History,
                        UserTurn = request.UserTurn,
                        UseNativeTools = false
                    };
                    return await SendOnceAsync(fallback, onDelta, ct).ConfigureAwait(false);
                }
                throw;
            }
        }

        private static bool IsNativeToolSchemaRejection(OmnixException ex)
        {
            if (ex == null || ex.Code != ErrorCode.PROVIDER_ERROR) return false;
            string details = ex.TechnicalDetails ?? "";
            return details.IndexOf("HTTP=400", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   details.IndexOf("HTTP=422", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private async Task<ChatResponse> SendOnceAsync(ChatRequest request, Action<string> onDelta, CancellationToken ct)
        {
            try
            {
                using (var client = HttpClientFactory.Create())
                using (var req = new HttpRequestMessage(HttpMethod.Post, BaseUrl + "/api/chat"))
                {
                    req.Content = new StringContent(BuildPayload(request), Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            string err = await SseLineReader.ReadErrorBodyAsync(response, ct).ConfigureAwait(false);
                            throw HttpStatusMapper.Map((int)response.StatusCode, err, "Ollama");
                        }

                        var sb = new StringBuilder();
                        var nativeCalls = new List<ProviderToolCall>();
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
                                    if (onDelta != null) onDelta(delta);
                                }

                                var toolCalls = obj.SelectToken("message.tool_calls") as JArray;
                                if (toolCalls != null)
                                {
                                    foreach (JObject toolCall in toolCalls.OfType<JObject>())
                                    {
                                        var decoded = DecodeOllamaToolCall(toolCall);
                                        if (decoded != null &&
                                            !nativeCalls.Any(x => string.Equals(x.Name, decoded.Name, StringComparison.Ordinal) &&
                                                                  string.Equals(x.ArgumentsJson, decoded.ArgumentsJson, StringComparison.Ordinal)))
                                            nativeCalls.Add(decoded);
                                    }
                                }

                                if ((bool?)obj["done"] == true) break;
                            }
                        }

                        if (nativeCalls.Count > 0)
                            Logger.Gateway("Ollama response contains native tool call(s): count=" + nativeCalls.Count);

                        return new ChatResponse
                        {
                            Text = sb.ToString(),
                            Model = Model,
                            ToolCalls = nativeCalls.Count > 0 ? nativeCalls : null
                        };
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
