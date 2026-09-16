using System;
using System.Collections.Generic;

namespace OMNIX.Core.Settings
{
    /// <summary>Layer 7.5 privacy modes. Default on first install: AskBeforeSending (most conservative).</summary>
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
        public string Name { get; set; }
        public string BaseUrl { get; set; }
        public string Model { get; set; }
        /// <summary>Result of the automatic Vision probe performed by Test Connection (null = unknown).</summary>
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
        public CustomProviderConfig CustomProvider { get; set; }
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
            s.Privacy = PrivacyMode.AskBeforeSending;
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
                { "custom", "gpt-4o-mini" }
            };
            s.CustomProvider = new CustomProviderConfig
            {
                Name = "My Endpoint",
                BaseUrl = "http://localhost:8080/v1",
                Model = "gpt-4o-mini"
            };
            s.HistoryMaxMessages = 500;
            s.HistoryMaxAgeDays = 30;
            s.ContextMaxCells = 2000;
            s.ContextMaxChars = 6000;
            s.ContextMaxTokens = 3000;
            s.PreferLocalWhenAvailable = true;
            return s;
        }
    }
}
