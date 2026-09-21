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
using OMNIX.Core.Errors;
using OMNIX.Core.Storage;
using OMNIX.Core.Tools;
using OMNIX.Core.Logging;

namespace OMNIX.Core.AiGateway.Http
{
    /// <summary>
    /// Shared OpenAI-compatible chat-completions client (stream: true, SSE).
    /// Used by Groq, OpenRouter, LM Studio and Custom providers so behavior is identical
    /// across them. Request history plus response/model bodies are bounded before full
    /// materialization so a malformed endpoint or long chat cannot consume unbounded memory
    /// inside Excel/Word/PowerPoint.
    /// </summary>
    public sealed class OpenAiCompatibleClient
    {
        private const int MaxAssistantChars = 2 * 1024 * 1024;
        private const int MaxJsonBodyBytes = 8 * 1024 * 1024;
        private const int MaxImageBytes = 20 * 1024 * 1024;
        private const int MaxModelCount = 5000;

        private readonly string _baseUrl;
        private readonly bool _anthropic;
        private readonly string _providerDisplayName;
        private readonly Dictionary<string, string> _extraHeaders;

        public OpenAiCompatibleClient(string baseUrl, string providerDisplayName, Dictionary<string, string> extraHeaders = null, bool anthropic = false)
        {
            if (string.IsNullOrWhiteSpace(baseUrl)) throw new ArgumentException("baseUrl is required", "baseUrl");
            _baseUrl = baseUrl.TrimEnd('/');
            _providerDisplayName = providerDisplayName;
            _extraHeaders = extraHeaders;
            _anthropic = anthropic;
        }

        private static JObject BuildMessage(ChatTurn turn)
        {
            var msg = new JObject();
            msg["role"] = RoleToString(turn.Role);

            if (turn.HasImages)
            {
                var content = new JArray();
                if (!string.IsNullOrEmpty(turn.Text))
                    content.Add(new JObject { { "type", "text" }, { "text", turn.Text } });
                foreach (var img in turn.Images.Where(i => i != null && i.PngBytes != null && i.PngBytes.Length > 0))
                {
                    if (img.PngBytes.Length > MaxImageBytes)
                        throw OmnixException.Model("Image attachment exceeds the 20 MB OMNIX safety limit.");
                    string b64 = Convert.ToBase64String(img.PngBytes);
                    content.Add(new JObject
                    {
                        { "type", "image_url" },
                        { "image_url", new JObject { { "url", "data:image/png;base64," + b64 } } }
                    });
                }
                msg["content"] = content;
            }
            else
            {
                msg["content"] = turn.Text ?? "";
            }
            return msg;
        }

        private static string RoleToString(ChatRole role)
        {
            switch (role)
            {
                case ChatRole.Assistant: return "assistant";
                case ChatRole.System: return "system";
                default: return "user";
            }
        }

        private static JObject BuildNativeToolParameters()
        {
            var names = new JArray();
            foreach (string name in ToolNames.AllWhitelisted) names.Add(name);
            return new JObject
            {
                { "type", "object" },
                { "additionalProperties", false },
                { "properties", new JObject
                    {
                        { "tool", new JObject
                            {
                                { "type", "string" },
                                { "enum", names },
                                { "description", "Exact OMNIX tool name from the hard whitelist." }
                            }
                        },
                        { "args", new JObject
                            {
                                { "type", "object" },
                                { "additionalProperties", true },
                                { "description", "Arguments for the selected OMNIX tool." }
                            }
                        }
                    }
                },
                { "required", new JArray("tool", "args") }
            };
        }

        private static JObject BuildOpenAiNativeTool()
        {
            return new JObject
            {
                { "type", "function" },
                { "function", new JObject
                    {
                        { "name", "omnix_tool" },
                        { "description", "Execute exactly one approved OMNIX Office tool. Use one tool call per provider turn." },
                        { "parameters", BuildNativeToolParameters() }
                    }
                }
            };
        }

        private static JObject BuildAnthropicNativeTool()
        {
            return new JObject
            {
                { "name", "omnix_tool" },
                { "description", "Execute exactly one approved OMNIX Office tool. Use one tool call per provider turn." },
                { "input_schema", BuildNativeToolParameters() }
            };
        }

        private static ProviderToolCall DecodeProviderToolCall(string id, string functionName, string argumentsJson)
        {
            string function = (functionName ?? "").Trim();
            if (string.Equals(function, "omnix_tool", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var root = string.IsNullOrWhiteSpace(argumentsJson) ? new JObject() : JObject.Parse(argumentsJson);
                    string tool = ToolNames.Normalize((string)root["tool"] ?? "");
                    JToken args = root["args"];
                    string argsJson = args == null ? "{}" :
                        args.Type == JTokenType.String ? (string)args : args.ToString(Formatting.None);
                    return new ProviderToolCall { Id = id ?? "", Name = tool, ArgumentsJson = string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson };
                }
                catch
                {
                    return new ProviderToolCall { Id = id ?? "", Name = "", ArgumentsJson = "__MALFORMED_NATIVE_TOOL_ARGUMENTS__" };
                }
            }

            // Compatibility: some OpenAI-like servers ignore the wrapper schema and emit the
            // canonical tool name directly. Accept it only if normalization lands on the existing
            // hard whitelist; the Gateway/ToolExecutor still validates it again.
            string normalized = ToolNames.Normalize(function);
            return new ProviderToolCall
            {
                Id = id ?? "",
                Name = normalized,
                ArgumentsJson = string.IsNullOrWhiteSpace(argumentsJson) ? "{}" : argumentsJson
            };
        }

        private sealed class StreamingToolAccumulator
        {
            public string Id;
            public string Name;
            public readonly StringBuilder Arguments = new StringBuilder();
        }

        private static List<ProviderToolCall> ParseOpenAiNonStreamingToolCalls(JObject root)
        {
            var result = new List<ProviderToolCall>();
            var calls = root.SelectToken("choices[0].message.tool_calls") as JArray;
            if (calls != null)
            {
                foreach (JObject call in calls.OfType<JObject>())
                {
                    string id = (string)call["id"] ?? "";
                    string name = (string)call.SelectToken("function.name") ?? "";
                    string args = (string)call.SelectToken("function.arguments") ?? "{}";
                    result.Add(DecodeProviderToolCall(id, name, args));
                }
            }
            else
            {
                var legacy = root.SelectToken("choices[0].message.function_call") as JObject;
                if (legacy != null)
                    result.Add(DecodeProviderToolCall("", (string)legacy["name"] ?? "", (string)legacy["arguments"] ?? "{}"));
            }
            return result;
        }

        private static List<ProviderToolCall> ParseAnthropicNonStreamingToolCalls(JObject root)
        {
            var result = new List<ProviderToolCall>();
            foreach (JObject part in (root["content"] as JArray ?? new JArray()).OfType<JObject>())
            {
                if (!string.Equals((string)part["type"], "tool_use", StringComparison.OrdinalIgnoreCase)) continue;
                string id = (string)part["id"] ?? "";
                string name = (string)part["name"] ?? "";
                string args = part["input"] != null ? part["input"].ToString(Formatting.None) : "{}";
                result.Add(DecodeProviderToolCall(id, name, args));
            }
            return result;
        }

        private static List<ProviderToolCall> MaterializeStreamingCalls(
            Dictionary<int, StreamingToolAccumulator> accumulators)
        {
            var result = new List<ProviderToolCall>();
            foreach (var kv in accumulators.OrderBy(x => x.Key))
            {
                var a = kv.Value;
                result.Add(DecodeProviderToolCall(a.Id, a.Name, a.Arguments.ToString()));
            }
            return result;
        }

        private string _configuredModel;

        public void SetModel(string model) { _configuredModel = model; }

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
                { "model", _configuredModel ?? "default" },
                { "messages", messages },
                { "stream", stream }
            };

            if (request.UseNativeTools)
            {
                if (_anthropic)
                {
                    payload["tools"] = new JArray(BuildAnthropicNativeTool());
                    payload["tool_choice"] = new JObject { { "type", "auto" } };
                }
                else
                {
                    payload["tools"] = new JArray(BuildOpenAiNativeTool());
                    payload["tool_choice"] = "auto";
                }
            }

            if (_anthropic)
            {
                payload["max_tokens"] = 4096;
                if (!string.IsNullOrEmpty(request.SystemPrompt)) payload["system"] = request.SystemPrompt;
                var converted = new JArray();
                foreach (JObject message in messages)
                {
                    if ((string)message["role"] == "system") continue;
                    var parts = message["content"] as JArray;
                    if (parts != null)
                        foreach (JObject part in parts)
                            if ((string)part["type"] == "image_url")
                            {
                                string url = (string)part.SelectToken("image_url.url");
                                part.Remove("image_url");
                                part["type"] = "image";
                                part["source"] = new JObject { { "type", "base64" }, { "media_type", "image/png" }, { "data", url.Substring(url.IndexOf(',') + 1) } };
                            }
                    converted.Add(message);
                }
                payload["messages"] = converted;
            }
            return payload.ToString(Formatting.None);
        }

        public async Task<ChatResponse> SendAsync(ChatRequest request, string apiKey, string model,
            Action<string> onDelta, CancellationToken ct)
        {
            _configuredModel = model;
            bool stream = onDelta != null;
            try
            {
                string body = BuildPayload(request, stream);
                return await SendRawAsync(body, apiKey, onDelta, ct).ConfigureAwait(false);
            }
            catch (OmnixException ex)
            {
                if (request != null && request.UseNativeTools && IsNativeToolSchemaRejection(ex))
                {
                    Logger.Gateway("Provider rejected native tool schema; retrying this provider turn with OMNIX text-protocol fallback. provider=" +
                                   _providerDisplayName + "; category=native_tools_unsupported");
                    var fallback = new ChatRequest
                    {
                        SystemPrompt = request.SystemPrompt,
                        History = request.History,
                        UserTurn = request.UserTurn,
                        UseNativeTools = false
                    };
                    string body = BuildPayload(fallback, stream);
                    return await SendRawAsync(body, apiKey, onDelta, ct).ConfigureAwait(false);
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

        public async Task<ChatResponse> SendRawAsync(string jsonBody, string apiKey,
            Action<string> onDelta, CancellationToken ct)
        {
            try
            {
                using (var client = HttpClientFactory.Create())
                using (var req = new HttpRequestMessage(HttpMethod.Post, _baseUrl + (_anthropic ? "/messages" : "/chat/completions")))
                {
                    if (_anthropic)
                    {
                        req.Headers.TryAddWithoutValidation("anthropic-version", "2023-06-01");
                        if (!string.IsNullOrEmpty(apiKey)) req.Headers.TryAddWithoutValidation("x-api-key", apiKey);
                    }
                    else if (!string.IsNullOrEmpty(apiKey))
                        req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                    if (_extraHeaders != null)
                        foreach (var kv in _extraHeaders)
                            req.Headers.TryAddWithoutValidation(kv.Key, kv.Value);

                    req.Content = new StringContent(jsonBody ?? "{}", Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            string err = await SseLineReader.ReadErrorBodyAsync(response, ct).ConfigureAwait(false);
                            throw HttpStatusMapper.Map((int)response.StatusCode, err, _providerDisplayName);
                        }

                        var sb = new StringBuilder();
                        var nativeCalls = new List<ProviderToolCall>();
                        string model = _configuredModel;

                        if (onDelta == null)
                        {
                            string full = await ReadBodyBoundedAsync(response.Content, MaxJsonBodyBytes, ct).ConfigureAwait(false);
                            var root = JObject.Parse(full);
                            model = (string)root.SelectToken("model") ?? model;
                            string text;
                            if (_anthropic)
                            {
                                text = string.Concat((root["content"] as JArray ?? new JArray())
                                    .Where(x => (string)x["type"] == "text")
                                    .Select(x => (string)x["text"]));
                                nativeCalls.AddRange(ParseAnthropicNonStreamingToolCalls(root));
                            }
                            else
                            {
                                text = (string)root.SelectToken("choices[0].message.content") ?? "";
                                nativeCalls.AddRange(ParseOpenAiNonStreamingToolCalls(root));
                            }

                            if (text.Length > MaxAssistantChars)
                                throw OmnixException.Provider(_providerDisplayName + " returned an over-sized assistant response.");
                            sb.Append(text);
                        }
                        else
                        {
                            var accumulators = new Dictionary<int, StreamingToolAccumulator>();
                            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                            {
                                foreach (string data in SseLineReader.ReadDataLines(stream, ct))
                                {
                                    if (data == "[DONE]") break;
                                    JObject chunk;
                                    try { chunk = JObject.Parse(data); }
                                    catch { continue; }

                                    model = (string)chunk.SelectToken("model") ?? model;
                                    if ((string)chunk["type"] == "error")
                                        throw OmnixException.Provider("Provider streaming error; response body redacted.");

                                    if (_anthropic)
                                    {
                                        string type = (string)chunk["type"] ?? "";
                                        int index = (int?)chunk["index"] ?? 0;
                                        if (type == "content_block_start" &&
                                            string.Equals((string)chunk.SelectToken("content_block.type"), "tool_use", StringComparison.OrdinalIgnoreCase))
                                        {
                                            StreamingToolAccumulator acc;
                                            if (!accumulators.TryGetValue(index, out acc))
                                            {
                                                acc = new StreamingToolAccumulator();
                                                accumulators[index] = acc;
                                            }
                                            acc.Id = (string)chunk.SelectToken("content_block.id") ?? acc.Id;
                                            acc.Name = (string)chunk.SelectToken("content_block.name") ?? acc.Name;
                                            var input = chunk.SelectToken("content_block.input") as JObject;
                                            if (input != null && input.Count > 0 && acc.Arguments.Length == 0)
                                                acc.Arguments.Append(input.ToString(Formatting.None));
                                        }
                                        else if (type == "content_block_delta" &&
                                                 string.Equals((string)chunk.SelectToken("delta.type"), "input_json_delta", StringComparison.OrdinalIgnoreCase))
                                        {
                                            StreamingToolAccumulator acc;
                                            if (!accumulators.TryGetValue(index, out acc))
                                            {
                                                acc = new StreamingToolAccumulator();
                                                accumulators[index] = acc;
                                            }
                                            string partial = (string)chunk.SelectToken("delta.partial_json");
                                            if (!string.IsNullOrEmpty(partial)) acc.Arguments.Append(partial);
                                        }

                                        string delta = (string)chunk.SelectToken("delta.text");
                                        if (!string.IsNullOrEmpty(delta))
                                        {
                                            if (sb.Length + delta.Length > MaxAssistantChars)
                                                throw OmnixException.Provider(_providerDisplayName + " streamed an over-sized assistant response.");
                                            sb.Append(delta);
                                            onDelta(delta);
                                        }
                                    }
                                    else
                                    {
                                        var toolDeltas = chunk.SelectToken("choices[0].delta.tool_calls") as JArray;
                                        if (toolDeltas != null)
                                        {
                                            foreach (JObject toolDelta in toolDeltas.OfType<JObject>())
                                            {
                                                int index = (int?)toolDelta["index"] ?? 0;
                                                StreamingToolAccumulator acc;
                                                if (!accumulators.TryGetValue(index, out acc))
                                                {
                                                    acc = new StreamingToolAccumulator();
                                                    accumulators[index] = acc;
                                                }
                                                string id = (string)toolDelta["id"];
                                                string name = (string)toolDelta.SelectToken("function.name");
                                                string args = (string)toolDelta.SelectToken("function.arguments");
                                                if (!string.IsNullOrEmpty(id)) acc.Id = id;
                                                if (!string.IsNullOrEmpty(name)) acc.Name = name;
                                                if (!string.IsNullOrEmpty(args)) acc.Arguments.Append(args);
                                            }
                                        }
                                        else
                                        {
                                            var legacy = chunk.SelectToken("choices[0].delta.function_call") as JObject;
                                            if (legacy != null)
                                            {
                                                StreamingToolAccumulator acc;
                                                if (!accumulators.TryGetValue(0, out acc))
                                                {
                                                    acc = new StreamingToolAccumulator();
                                                    accumulators[0] = acc;
                                                }
                                                string name = (string)legacy["name"];
                                                string args = (string)legacy["arguments"];
                                                if (!string.IsNullOrEmpty(name)) acc.Name = name;
                                                if (!string.IsNullOrEmpty(args)) acc.Arguments.Append(args);
                                            }
                                        }

                                        string delta = (string)chunk.SelectToken("choices[0].delta.content");
                                        if (!string.IsNullOrEmpty(delta))
                                        {
                                            if (sb.Length + delta.Length > MaxAssistantChars)
                                                throw OmnixException.Provider(_providerDisplayName + " streamed an over-sized assistant response.");
                                            sb.Append(delta);
                                            onDelta(delta);
                                        }
                                    }
                                }
                            }
                            nativeCalls.AddRange(MaterializeStreamingCalls(accumulators));
                        }

                        if (nativeCalls.Count > 0)
                            Logger.Gateway("Provider response contains native tool call(s): count=" + nativeCalls.Count +
                                           "; provider=" + _providerDisplayName);

                        return new ChatResponse
                        {
                            Text = sb.ToString(),
                            Model = model,
                            ToolCalls = nativeCalls.Count > 0 ? nativeCalls : null
                        };
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (HttpRequestException ex)
            {
                throw OmnixException.Network(_providerDisplayName + ": " + ex.Message);
            }
            catch (OmnixException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw OmnixException.Provider(_providerDisplayName + " transport failure: " + ex.Message);
            }
        }

        public async Task<IReadOnlyList<string>> ListModelsAsync(string apiKey, CancellationToken ct)
        {
            try
            {
                var list = new List<string>();
                var cursors = new HashSet<string>(StringComparer.Ordinal);
                string cursor = null;
                using (var client = HttpClientFactory.Create(TimeSpan.FromSeconds(20)))
                for (int page = 0; page < 50; page++)
                {
                    string url = _baseUrl + "/models" + (cursor == null ? "" : "?after_id=" + Uri.EscapeDataString(cursor));
                    using (var req = new HttpRequestMessage(HttpMethod.Get, url))
                    {
                        if (_anthropic)
                        {
                            req.Headers.TryAddWithoutValidation("anthropic-version", "2023-06-01");
                            if (!string.IsNullOrEmpty(apiKey)) req.Headers.TryAddWithoutValidation("x-api-key", apiKey);
                        }
                        else if (!string.IsNullOrEmpty(apiKey))
                            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                        if (_extraHeaders != null)
                            foreach (var kv in _extraHeaders) req.Headers.TryAddWithoutValidation(kv.Key, kv.Value);
                        using (var response = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                        {
                            if (!response.IsSuccessStatusCode)
                            {
                                string err = await SseLineReader.ReadErrorBodyAsync(response, ct).ConfigureAwait(false);
                                throw HttpStatusMapper.Map((int)response.StatusCode, err, _providerDisplayName);
                            }
                            string json = await ReadBodyBoundedAsync(response.Content, MaxJsonBodyBytes, ct).ConfigureAwait(false);
                            var root = JObject.Parse(json);
                            foreach (var m in root["data"] ?? new JArray())
                            {
                                string id = (string)m["id"];
                                if (!string.IsNullOrWhiteSpace(id) && !list.Contains(id)) list.Add(id);
                                if (list.Count >= MaxModelCount) return list;
                            }
                            if (!_anthropic || (bool?)root["has_more"] != true) return list;
                            cursor = (string)root["last_id"];
                            if (string.IsNullOrEmpty(cursor) || !cursors.Add(cursor))
                                throw OmnixException.Provider("Model catalog returned an invalid pagination cursor. Enter a model ID manually.");
                        }
                    }
                }
                return list;
            }
            catch (OperationCanceledException) { throw; }
            catch (OmnixException) { throw; }
            catch (HttpRequestException) { throw OmnixException.Network(_providerDisplayName + " model discovery could not reach the endpoint."); }
            catch (Exception) { throw OmnixException.Provider(_providerDisplayName + " returned an invalid model catalog. Enter a model ID manually."); }
        }

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
                    if (remaining <= 0) throw new InvalidDataException("HTTP response body exceeded OMNIX safety limit.");
                    int read = await stream.ReadAsync(buffer, 0, Math.Min(buffer.Length, remaining), ct).ConfigureAwait(false);
                    if (read <= 0) break;
                    ms.Write(buffer, 0, read);
                    if (ms.Length > maxBytes) throw new InvalidDataException("HTTP response body exceeded OMNIX safety limit.");
                }
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
    }
}
