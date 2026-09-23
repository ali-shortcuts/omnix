using System;
using System.Collections.Generic;

namespace OMNIX.Core.Settings
{
    /// <summary>Layer 7.5 privacy modes. Default on first install: CloudAllowed; existing explicit choices are preserved.</summary>
    public enum PrivacyMode
    {
        LocalOnly = 0,
        CloudAllowed = 1,
        AskBeforeSending = 2
    }

    public enum ThemeMode
    {
        System = 0,
        Light = 1,
        Dark = 2
    }

    public sealed class CustomProviderConfig
    {
        public string ApiType { get; set; } // OpenAI (default for existing settings) or Anthropic
        public string Name { get; set; }
        public string BaseUrl { get; set; }
        public string Model { get; set; }
        /// <summary>Last explicit model-capability probe result when available (null = unknown). Connection tests never infer Vision support.</summary>
        public bool? SupportsVision { get; set; }
    }

    /// <summary>
    /// POCO settings. API keys are NEVER stored here in plain text — SettingsManager keeps them
    /// DPAPI-protected in a separate dictionary.
    /// </summary>
    public sealed class OmnixSettings
    {
        public int SchemaVersion { get; set; }

        public PrivacyMode Privacy { get; set; }
        public ThemeMode Theme { get; set; }
        public string UiLanguage { get; set; }
        public string SelectedProviderId { get; set; }
        public string PreferredLocalProviderId { get; set; }
        public Dictionary<string, string> Models { get; set; }
        public Dictionary<string, List<string>> SavedModels { get; set; }
        public int ExecutionStepDelayMs { get; set; }
        public string BusinessLocale { get; set; }
        public string CloudflareAccountId { get; set; }
        public bool ConfirmEveryWrite { get; set; }
        public CustomProviderConfig CustomProvider { get; set; }
        public CustomProviderConfig AgentRouter { get; set; }
        public CustomProviderConfig EndpointConfig(string id)
        {
            if (id != "agentrouter") return CustomProvider;
            if (AgentRouter == null) AgentRouter = new CustomProviderConfig { Name = "Agent Router", BaseUrl = "https://agentrouter.org/v1", ApiType = "Anthropic", Model = "" };
            return AgentRouter;
        }
        public int HistoryMaxMessages { get; set; }
        public int HistoryMaxAgeDays { get; set; }
        public int ContextMaxCells { get; set; }
        public int ContextMaxChars { get; set; }
        public int ContextMaxTokens { get; set; }
        public bool PreferLocalWhenAvailable { get; set; }

        public static OmnixSettings CreateDefaults()
        {
            var s = new OmnixSettings();
            s.SchemaVersion = 4;
            s.Privacy = PrivacyMode.CloudAllowed;
            s.SavedModels = new Dictionary<string, List<string>>();
            s.ExecutionStepDelayMs = 350;
            s.BusinessLocale = "Afghanistan; Dari; currency AFN";
            s.Theme = ThemeMode.System;
            s.UiLanguage = "en";
            s.SelectedProviderId = "gemini";
            s.PreferredLocalProviderId = "ollama";
            s.Models = new Dictionary<string, string>
            {
                { "gemini", "" },
                { "groq", "openai/gpt-oss-120b" },
                { "openrouter", "openrouter/free" },
                { "mistral", "mistral-small-latest" },
                { "huggingface", "openai/gpt-oss-120b:fastest" },
                { "cerebras", "gpt-oss-120b" },
                { "ollama", "" },
                { "lmstudio", "" },
                { "custom", "" }
            };
            s.CustomProvider = new CustomProviderConfig
            {
                Name = "My Endpoint",
                BaseUrl = "http://localhost:8080/v1",
                Model = ""
            };
            s.HistoryMaxMessages = 500;
            s.HistoryMaxAgeDays = 30;
            s.ContextMaxCells = 2000;
            s.ContextMaxChars = 6000;
            s.ContextMaxTokens = 3000;
            s.PreferLocalWhenAvailable = false;
            return s;
        }
    }
}
