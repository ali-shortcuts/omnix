using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using OMNIX.Core.Logging;

namespace OMNIX.Core.Settings
{
    /// <summary>
    /// Layer 8 (Storage): settings + DPAPI-encrypted API keys at %LOCALAPPDATA%\OMNIX\settings.dat.
    /// API keys are never stored in plain text and never logged. Schema migrations preserve
    /// explicit user choices and only replace model ids that exactly match known obsolete OMNIX defaults.
    /// </summary>
    public sealed class SettingsManager
    {
        private static readonly object Gate = new object();
        private static SettingsManager _instance;
        public static SettingsManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (Gate)
                    {
                        if (_instance == null) _instance = new SettingsManager();
                    }
                }
                return _instance;
            }
        }

        private static readonly byte[] Magic = Encoding.ASCII.GetBytes("OMNIXS01");
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("OMNIX::v1::DPAPI");

        private readonly string _path;
        private readonly Dictionary<string, string> _plainKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public OmnixSettings Settings { get; private set; }

        public SettingsManager()
        {
            _path = Path.Combine(Logger.BaseDir, "settings.dat");
            Load();
        }

        public string SettingsFilePath { get { return _path; } }

        private void Load()
        {
            try
            {
                if (!File.Exists(_path))
                {
                    Settings = OmnixSettings.CreateDefaults();
                    Logger.Startup("settings.dat not found — defaults created (privacy=AskBeforeSending)");
                    return;
                }

                byte[] blob = File.ReadAllBytes(_path);
                if (blob.Length < Magic.Length + 4 || !StartsWith(blob, Magic))
                {
                    BackupCorrupt("bad magic");
                    Settings = OmnixSettings.CreateDefaults();
                    return;
                }

                string json = Encoding.UTF8.GetString(blob, Magic.Length, blob.Length - Magic.Length);
                var dto = JsonConvert.DeserializeObject<SettingsDto>(json);
                if (dto == null || dto.Settings == null)
                {
                    BackupCorrupt("null payload");
                    Settings = OmnixSettings.CreateDefaults();
                    return;
                }

                Settings = dto.Settings;
                bool migrated = MigrateSettingsIfNeeded();

                lock (_plainKeys)
                {
                    _plainKeys.Clear();
                    if (dto.ProtectedKeys != null)
                    {
                        foreach (var kv in dto.ProtectedKeys)
                        {
                            try
                            {
                                byte[] prot = Convert.FromBase64String(kv.Value);
                                byte[] plain = ProtectedData.Unprotect(prot, Entropy, DataProtectionScope.CurrentUser);
                                try
                                {
                                    _plainKeys[kv.Key] = Encoding.UTF8.GetString(plain);
                                }
                                finally
                                {
                                    Array.Clear(plain, 0, plain.Length);
                                    Array.Clear(prot, 0, prot.Length);
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Error("settings", "Could not unprotect API key for provider '" + kv.Key + "' — key reset.", ex);
                            }
                        }
                    }
                }

                if (migrated)
                {
                    Save();
                    Logger.Startup("settings schema migrated to v" + Settings.SchemaVersion + " without replacing explicit user model choices");
                }

                Logger.Startup("settings loaded: provider=" + Settings.SelectedProviderId + " privacy=" + Settings.Privacy);
            }
            catch (Exception ex)
            {
                Logger.Error("settings", "Failed to load settings — defaults used.", ex);
                Settings = OmnixSettings.CreateDefaults();
            }
        }

        private bool MigrateSettingsIfNeeded()
        {
            bool changed = false;
            if (Settings == null) return false;

            if (Settings.Models == null)
            {
                Settings.Models = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                changed = true;
            }

            // v1 -> v2: replace only the exact cloud defaults previously shipped by OMNIX.
            if (Settings.SchemaVersion < 2)
            {
                string model;
                if (!Settings.Models.TryGetValue("gemini", out model) || string.IsNullOrWhiteSpace(model) ||
                    string.Equals(model, "gemini-2.0-flash", StringComparison.OrdinalIgnoreCase))
                    Settings.Models["gemini"] = "gemini-3.6-flash";

                if (!Settings.Models.TryGetValue("groq", out model) || string.IsNullOrWhiteSpace(model) ||
                    string.Equals(model, "llama-3.3-70b-versatile", StringComparison.OrdinalIgnoreCase))
                    Settings.Models["groq"] = "openai/gpt-oss-120b";

                if (!Settings.Models.ContainsKey("openrouter") || string.IsNullOrWhiteSpace(Settings.Models["openrouter"]))
                    Settings.Models["openrouter"] = "openrouter/auto";

                EnsureCommonDefaults();
                Settings.SchemaVersion = 2;
                changed = true;
            }

            // v2 -> v3: add researched cloud providers and update OMNIX-owned old defaults only.
            if (Settings.SchemaVersion < 3)
            {
                string model;
                if (!Settings.Models.TryGetValue("gemini", out model) || string.IsNullOrWhiteSpace(model) ||
                    string.Equals(model, "gemini-3.6-flash", StringComparison.OrdinalIgnoreCase))
                    Settings.Models["gemini"] = "gemini-3.8-flash";

                if (!Settings.Models.TryGetValue("openrouter", out model) || string.IsNullOrWhiteSpace(model) ||
                    string.Equals(model, "openrouter/auto", StringComparison.OrdinalIgnoreCase))
                    Settings.Models["openrouter"] = "openrouter/free";

                if (!Settings.Models.ContainsKey("mistral"))
                    Settings.Models["mistral"] = "mistral-small-latest";
                if (!Settings.Models.ContainsKey("cerebras"))
                    Settings.Models["cerebras"] = "gpt-oss-120b";

                EnsureCommonDefaults();
                Settings.SchemaVersion = 3;
                changed = true;
            }

            // v3 -> v4: add Hugging Face Inference Providers. Existing users keep every prior
            // provider/model/API-key choice; this migration only introduces the new model slot.
            if (Settings.SchemaVersion < 4)
            {
                if (!Settings.Models.ContainsKey("huggingface") ||
                    string.IsNullOrWhiteSpace(Settings.Models["huggingface"]))
                    Settings.Models["huggingface"] = "openai/gpt-oss-120b:fastest";

                EnsureCommonDefaults();
                Settings.SchemaVersion = 4;
                changed = true;
            }

            return changed;
        }

        private void EnsureCommonDefaults()
        {
            if (!Settings.Models.ContainsKey("gemini")) Settings.Models["gemini"] = "gemini-3.8-flash";
            if (!Settings.Models.ContainsKey("groq")) Settings.Models["groq"] = "openai/gpt-oss-120b";
            if (!Settings.Models.ContainsKey("openrouter")) Settings.Models["openrouter"] = "openrouter/free";
            if (!Settings.Models.ContainsKey("mistral")) Settings.Models["mistral"] = "mistral-small-latest";
            if (!Settings.Models.ContainsKey("huggingface")) Settings.Models["huggingface"] = "openai/gpt-oss-120b:fastest";
            if (!Settings.Models.ContainsKey("cerebras")) Settings.Models["cerebras"] = "gpt-oss-120b";
            if (!Settings.Models.ContainsKey("ollama")) Settings.Models["ollama"] = "";
            if (!Settings.Models.ContainsKey("lmstudio")) Settings.Models["lmstudio"] = "";
            if (!Settings.Models.ContainsKey("custom")) Settings.Models["custom"] = "gpt-4o-mini";

            if (string.IsNullOrWhiteSpace(Settings.PreferredLocalProviderId)) Settings.PreferredLocalProviderId = "ollama";
            if (string.IsNullOrWhiteSpace(Settings.SelectedProviderId)) Settings.SelectedProviderId = "gemini";
            if (string.IsNullOrWhiteSpace(Settings.UiLanguage)) Settings.UiLanguage = "en";

            if (Settings.CustomProvider == null)
            {
                Settings.CustomProvider = new CustomProviderConfig
                {
                    Name = "My Endpoint",
                    BaseUrl = "http://localhost:8080/v1",
                    Model = "gpt-4o-mini"
                };
            }

            if (Settings.HistoryMaxMessages <= 0) Settings.HistoryMaxMessages = 500;
            if (Settings.HistoryMaxAgeDays <= 0) Settings.HistoryMaxAgeDays = 30;
            if (Settings.ContextMaxCells <= 0) Settings.ContextMaxCells = 2000;
            if (Settings.ContextMaxChars <= 0) Settings.ContextMaxChars = 6000;
            if (Settings.ContextMaxTokens <= 0) Settings.ContextMaxTokens = 3000;
        }

        public bool LastSaveSucceeded { get; private set; }

        public void Save()
        {
            LastSaveSucceeded = false;
            string tmp = null;
            var sensitiveBuffers = new List<byte[]>();
            try
            {
                var dto = new SettingsDto { Settings = Settings };

                lock (_plainKeys)
                {
                    dto.ProtectedKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kv in _plainKeys)
                    {
                        byte[] plain = Encoding.UTF8.GetBytes(kv.Value ?? "");
                        byte[] prot = null;
                        sensitiveBuffers.Add(plain);
                        try
                        {
                            prot = ProtectedData.Protect(plain, Entropy, DataProtectionScope.CurrentUser);
                            dto.ProtectedKeys[kv.Key] = Convert.ToBase64String(prot);
                        }
                        finally
                        {
                            if (prot != null) Array.Clear(prot, 0, prot.Length);
                        }
                    }
                }

                string json = JsonConvert.SerializeObject(dto, Formatting.Indented);
                Directory.CreateDirectory(Path.GetDirectoryName(_path));
                tmp = _path + ".tmp";
                byte[] fileBytes = Concat(Magic, Encoding.UTF8.GetBytes(json));
                try
                {
                    using (var stream = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        stream.Write(fileBytes, 0, fileBytes.Length);
                        stream.Flush(true);
                    }
                }
                finally
                {
                    Array.Clear(fileBytes, 0, fileBytes.Length);
                }

                if (File.Exists(_path))
                {
                    // Same-volume atomic replacement avoids the old delete-then-move window where
                    // an Office/process crash could leave the user with no settings.dat at all.
                    File.Replace(tmp, _path, null);
                }
                else
                {
                    File.Move(tmp, _path);
                }
                tmp = null;
                LastSaveSucceeded = true;
            }
            catch (Exception ex)
            {
                Logger.Error("settings", "Failed to save settings.", ex);
            }
            finally
            {
                foreach (var buffer in sensitiveBuffers)
                {
                    if (buffer != null) Array.Clear(buffer, 0, buffer.Length);
                }
                if (!string.IsNullOrEmpty(tmp))
                {
                    try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                }
            }
        }

        public string GetApiKey(string providerId)
        {
            if (string.IsNullOrEmpty(providerId)) return null;
            lock (_plainKeys)
            {
                string v;
                return _plainKeys.TryGetValue(providerId, out v) ? v : null;
            }
        }

        public void SetApiKey(string providerId, string plainKey)
        {
            lock (_plainKeys)
            {
                if (string.IsNullOrEmpty(plainKey)) _plainKeys.Remove(providerId);
                else _plainKeys[providerId] = plainKey;
            }
            Save();
        }

        public bool HasApiKey(string providerId)
        {
            return !string.IsNullOrEmpty(GetApiKey(providerId));
        }

        private void BackupCorrupt(string reason)
        {
            try
            {
                if (File.Exists(_path))
                {
                    string bak = _path + ".corrupt-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
                    File.Copy(_path, bak, true);
                    Logger.Startup("settings.dat corrupt (" + reason + ") — backed up to " + bak + " and reset to defaults.");
                }
            }
            catch { }
        }

        private static bool StartsWith(byte[] blob, byte[] prefix)
        {
            for (int i = 0; i < prefix.Length; i++)
                if (blob[i] != prefix[i]) return false;
            return true;
        }

        private static byte[] Concat(byte[] a, byte[] b)
        {
            var r = new byte[a.Length + b.Length];
            Buffer.BlockCopy(a, 0, r, 0, a.Length);
            Buffer.BlockCopy(b, 0, r, a.Length, b.Length);
            return r;
        }

        private sealed class SettingsDto
        {
            [JsonProperty("settings")]
            public OmnixSettings Settings { get; set; }

            [JsonProperty("protectedKeys")]
            public Dictionary<string, string> ProtectedKeys { get; set; }
        }
    }
}
