using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OMNIX.Core.Errors;
using OMNIX.Core.Settings;
using OMNIX.Core.Storage;

namespace OMNIX.Core.AiGateway
{
    public sealed class ProviderCredentials
    {
        public string ApiKey { get; set; }
        public string Model { get; set; }
        public string BaseUrl { get; set; }
    }

    public sealed class ChatRequest
    {
        public string SystemPrompt { get; set; }
        public List<ChatTurn> History { get; set; }
        public ChatTurn UserTurn { get; set; }

        /// <summary>
        /// True only for the Office execution loop. Diagnostics/connection tests leave this false,
        /// so synthetic provider checks cannot unexpectedly invoke Office tools.
        /// </summary>
        public bool UseNativeTools { get; set; }

        public bool HasImages
        {
            get
            {
                if (UserTurn != null && UserTurn.HasImages) return true;
                if (History != null)
                    foreach (var t in History)
                        if (t != null && t.HasImages) return true;
                return false;
            }
        }
    }

    public sealed class ProviderToolCall
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ArgumentsJson { get; set; }
    }

    public sealed class ChatResponse
    {
        public string Text { get; set; }
        public string Model { get; set; }
        public bool WasCancelled { get; set; }
        public List<ProviderToolCall> ToolCalls { get; set; }

        public bool HasToolCalls
        {
            get { return ToolCalls != null && ToolCalls.Count > 0; }
        }
    }

    /// <summary>
    /// Hard provider-boundary budget plus deterministic conversation-continuity selection.
    /// Local history storage can retain more messages for UI/history, but every provider request is
    /// rebuilt through this limiter before routing. This prevents a long Office conversation, old
    /// screenshots or one giant message from expanding a request without bound inside
    /// Excel/Word/PowerPoint.
    ///
    /// The continuity selector is deliberately local and deterministic: it keeps a bounded recent
    /// tail, preserves the first meaningful user goal as an anchor when possible, and uses lexical
    /// overlap with the current turn to retain useful older turns. It never makes a second AI call,
    /// never persists a derived memory, and never increases the existing hard provider budget.
    ///
    /// Historical image bytes are intentionally NOT replayed. The newest/current user turn (and
    /// current tool-result turn) may contain bounded images; older image turns receive an explicit
    /// text marker so the model cannot pretend it still sees an image that was omitted.
    /// </summary>
    public static class ChatRequestBudgeter
    {
        private const int HardMaxHistoryTurns = 80;
        private const int HardMaxHistoryChars = 48 * 1024;
        private const int MaxSingleHistoryTurnChars = 12 * 1024;
        private const int MaxSystemPromptChars = 32 * 1024;
        private const int MaxCurrentTurnChars = 64 * 1024;
        private const int MaxCurrentImages = 4;
        private const int MaxImageBytes = 20 * 1024 * 1024;
        private const int MaxCurrentImageBytesTotal = 24 * 1024 * 1024;
        private const int RecentHistoryTurns = 20;
        private const int MaxContinuityTerms = 48;
        private const string HistoricalImageMarker = "\n[OMNIX: an earlier image was omitted from provider replay to keep request memory bounded; ask the user to reattach it if visual inspection is required.]";

        private static readonly HashSet<string> ContinuityStopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "the", "and", "for", "with", "this", "that", "from", "have", "has", "had",
            "you", "your", "are", "was", "were", "will", "would", "could", "should", "can",
            "please", "continue", "then", "into", "about", "what", "when", "where", "which",
            "just", "more", "less", "than", "also", "only", "make", "need", "want", "using",
            "use", "its", "our", "their", "them", "they", "there", "here", "been", "being"
        };

        public static ChatRequest Apply(ChatRequest source)
        {
            if (source == null)
                throw new OmnixException(ErrorCode.CORE_ERROR,
                    "The AI request is empty.", "ChatRequestBudgeter.Apply received null.", "Retry the request.");
            if (source.UserTurn == null)
                throw new OmnixException(ErrorCode.CORE_ERROR,
                    "The AI request has no current user/tool turn.", "ChatRequest.UserTurn is null.", "Retry the request.");

            var settings = SettingsManager.Instance.Settings;
            int configuredTurns = settings != null ? settings.HistoryMaxMessages : HardMaxHistoryTurns;
            int maxTurns = Math.Max(1, Math.Min(HardMaxHistoryTurns, configuredTurns));

            // ContextMaxTokens is a user-controlled context preference, not an exact tokenizer.
            // Use it only as a conservative character budget hint, then enforce hard caps.
            int configuredTokens = settings != null ? settings.ContextMaxTokens : 3000;
            long hintedChars = (long)Math.Max(1000, configuredTokens) * 8L;
            int maxHistoryChars = (int)Math.Max(8 * 1024L, Math.Min(HardMaxHistoryChars, hintedChars));

            var result = new ChatRequest
            {
                SystemPrompt = TruncatePreservingEnds(source.SystemPrompt, MaxSystemPromptChars),
                History = BuildBoundedHistory(source.History, maxTurns, maxHistoryChars,
                    source.UserTurn != null ? source.UserTurn.Text : null),
                UserTurn = CloneCurrentTurn(source.UserTurn),
                UseNativeTools = source.UseNativeTools
            };
            return result;
        }

        /// <summary>
        /// Builds a bounded, chronologically ordered replay set with three priorities:
        /// 1) a recent tail for conversational coherence,
        /// 2) the first user goal as a stable task anchor,
        /// 3) older turns lexically relevant to the current request.
        /// Any remaining budget is filled newest-first. This avoids the previous failure mode where
        /// a long session could entirely forget its original task or a relevant middle decision.
        /// </summary>
        private static List<ChatTurn> BuildBoundedHistory(
            List<ChatTurn> source,
            int maxTurns,
            int maxChars,
            string currentText)
        {
            var result = new List<ChatTurn>();
            if (source == null || source.Count == 0 || maxTurns <= 0 || maxChars <= 0) return result;

            var selected = new Dictionary<int, ChatTurn>();
            int usedChars = 0;

            // Keep room for at least one continuity anchor when maxTurns permits it. The character
            // reserve is bounded to 25% (max 12 KB), so continuity can never crowd out the recent tail.
            int continuityCharReserve = maxTurns > 1
                ? Math.Max(1024, Math.Min(12 * 1024, maxChars / 4))
                : 0;
            int recentCharBudget = Math.Max(1, maxChars - continuityCharReserve);
            int recentTurnTarget = Math.Min(RecentHistoryTurns, maxTurns > 1 ? maxTurns - 1 : maxTurns);

            // 1) Recent tail: newest-first selection, final output is sorted chronologically later.
            int recentChars = 0;
            for (int i = source.Count - 1; i >= 0 && selected.Count < recentTurnTarget; i--)
            {
                var turn = source[i];
                if (turn == null) continue;

                int remaining = recentCharBudget - recentChars;
                if (remaining <= 0) break;
                var clone = CloneHistoricalTurn(turn, Math.Min(MaxSingleHistoryTurnChars, remaining));
                if (clone == null) continue;

                selected[i] = clone;
                int chars = clone.Text != null ? clone.Text.Length : 0;
                recentChars += chars;
                usedChars += chars;
            }

            // 2) Stable first-user anchor. This is especially important for short follow-ups such as
            // "continue", where lexical relevance alone provides almost no signal.
            int firstUserIndex = FindFirstUserIndex(source);
            AddContinuityCandidate(source, firstUserIndex, selected, maxTurns, maxChars, ref usedChars,
                continuityCharReserve);

            // 3) Older relevant turns. Only score turns outside the recent tail/anchor, and ignore
            // internal tool-protocol artifacts so stale tool JSON cannot be promoted as "memory".
            var currentTerms = ExtractContinuityTerms(currentText);
            if (currentTerms.Count > 0 && selected.Count < maxTurns && usedChars < maxChars)
            {
                var ranked = new List<Tuple<int, int>>();
                for (int i = 0; i < source.Count; i++)
                {
                    if (selected.ContainsKey(i) || source[i] == null) continue;
                    if (ContainsInternalToolProtocol(source[i].Text)) continue;
                    int score = ScoreContinuityRelevance(source[i], currentTerms);
                    if (score > 0) ranked.Add(Tuple.Create(i, score));
                }

                foreach (var candidate in ranked
                    .OrderByDescending(x => x.Item2)
                    .ThenByDescending(x => x.Item1))
                {
                    if (selected.Count >= maxTurns || usedChars >= maxChars) break;
                    AddContinuityCandidate(source, candidate.Item1, selected, maxTurns, maxChars,
                        ref usedChars, maxChars - usedChars);
                }
            }

            // 4) Use any remaining budget newest-first. This preserves the old behavior as a fallback
            // without allowing it to erase the anchor/relevance guarantees above.
            for (int i = source.Count - 1; i >= 0 && selected.Count < maxTurns && usedChars < maxChars; i--)
            {
                if (selected.ContainsKey(i) || source[i] == null) continue;
                int remaining = maxChars - usedChars;
                var clone = CloneHistoricalTurn(source[i], Math.Min(MaxSingleHistoryTurnChars, remaining));
                if (clone == null) continue;
                selected[i] = clone;
                usedChars += clone.Text != null ? clone.Text.Length : 0;
            }

            foreach (var item in selected.OrderBy(x => x.Key))
                result.Add(item.Value);
            return result;
        }

        private static int FindFirstUserIndex(List<ChatTurn> source)
        {
            if (source == null) return -1;
            for (int i = 0; i < source.Count; i++)
            {
                var turn = source[i];
                if (turn != null && turn.Role == ChatRole.User &&
                    (!string.IsNullOrWhiteSpace(turn.Text) || turn.HasImages))
                    return i;
            }
            return -1;
        }

        private static void AddContinuityCandidate(
            List<ChatTurn> source,
            int index,
            Dictionary<int, ChatTurn> selected,
            int maxTurns,
            int maxChars,
            ref int usedChars,
            int localBudget)
        {
            if (source == null || index < 0 || index >= source.Count || selected.ContainsKey(index)) return;
            if (selected.Count >= maxTurns || usedChars >= maxChars || localBudget <= 0) return;

            int remaining = Math.Min(maxChars - usedChars, localBudget);
            if (remaining <= 0) return;
            var clone = CloneHistoricalTurn(source[index], Math.Min(MaxSingleHistoryTurnChars, remaining));
            if (clone == null) return;
            selected[index] = clone;
            usedChars += clone.Text != null ? clone.Text.Length : 0;
        }

        private static ChatTurn CloneHistoricalTurn(ChatTurn turn, int maxChars)
        {
            if (turn == null || maxChars <= 0) return null;
            string text = turn.Text ?? string.Empty;
            if (turn.HasImages) text += HistoricalImageMarker;
            text = TruncatePreservingEnds(text, maxChars);

            if (text.Length == 0 && !turn.HasImages) return null;
            return new ChatTurn
            {
                Role = turn.Role,
                Text = text,
                Images = null,
                TimestampUtc = turn.TimestampUtc
            };
        }

        private static int ScoreContinuityRelevance(ChatTurn turn, HashSet<string> currentTerms)
        {
            if (turn == null || currentTerms == null || currentTerms.Count == 0) return 0;
            var turnTerms = ExtractContinuityTerms(turn.Text);
            if (turnTerms.Count == 0) return 0;

            int overlap = 0;
            foreach (string term in currentTerms)
                if (turnTerms.Contains(term)) overlap++;
            if (overlap == 0) return 0;

            int score = overlap * 100;
            if (turn.Role == ChatRole.User) score += 25;
            if (!string.IsNullOrEmpty(turn.Text) && turn.Text.Length <= 2000) score += 5;
            return score;
        }

        private static HashSet<string> ExtractContinuityTerms(string text)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(text)) return result;

            var token = new StringBuilder();
            Action flush = delegate
            {
                if (token.Length >= 2)
                {
                    string value = token.ToString();
                    if (!ContinuityStopWords.Contains(value) && result.Count < MaxContinuityTerms)
                        result.Add(value);
                }
                token.Clear();
            };

            foreach (char c in text)
            {
                if (char.IsLetterOrDigit(c)) token.Append(char.ToLowerInvariant(c));
                else flush();
                if (result.Count >= MaxContinuityTerms) break;
            }
            if (result.Count < MaxContinuityTerms) flush();
            return result;
        }

        private static bool ContainsInternalToolProtocol(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            return text.IndexOf("<tool_call>", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   text.IndexOf("```omnix_tool", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   text.IndexOf("OMNIX TOOL RESULT:", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static ChatTurn CloneCurrentTurn(ChatTurn turn)
        {
            string text = turn.Text ?? string.Empty;
            if (text.Length > MaxCurrentTurnChars)
                throw OmnixException.Model(
                    "This message is too large for one OMNIX request (maximum 64 KB of text). Split it into smaller messages.");

            List<ImageAttachment> images = null;
            if (turn.Images != null && turn.Images.Count > 0)
            {
                images = new List<ImageAttachment>();
                long totalBytes = 0;
                foreach (var image in turn.Images.Where(i => i != null && i.PngBytes != null && i.PngBytes.Length > 0))
                {
                    if (images.Count >= MaxCurrentImages)
                        throw OmnixException.Model("Too many images in one OMNIX request (maximum 4).");
                    if (image.PngBytes.Length > MaxImageBytes)
                        throw OmnixException.Model("An image exceeds the 20 MB OMNIX safety limit.");
                    totalBytes += image.PngBytes.Length;
                    if (totalBytes > MaxCurrentImageBytesTotal)
                        throw OmnixException.Model("Images in this request exceed the 24 MB combined OMNIX safety limit.");

                    // Share the immutable request byte array instead of duplicating tens of MB in
                    // the Office process. Provider adapters never mutate attachment bytes.
                    images.Add(new ImageAttachment
                    {
                        FileName = SafeLabel(image.FileName, 160),
                        SourceLabel = SafeLabel(image.SourceLabel, 120),
                        PngBytes = image.PngBytes
                    });
                }
            }

            return new ChatTurn
            {
                Role = turn.Role,
                Text = text,
                Images = images,
                TimestampUtc = turn.TimestampUtc
            };
        }

        private static string TruncatePreservingEnds(string value, int maxChars)
        {
            if (string.IsNullOrEmpty(value) || maxChars <= 0) return string.Empty;
            if (value.Length <= maxChars) return value;
            if (maxChars < 96) return value.Substring(value.Length - maxChars, maxChars);

            const string marker = "\n…[OMNIX request history truncated]…\n";
            int available = Math.Max(1, maxChars - marker.Length);
            int head = available / 2;
            int tail = available - head;
            return value.Substring(0, head) + marker + value.Substring(value.Length - tail, tail);
        }

        private static string SafeLabel(string value, int maxChars)
        {
            value = value ?? string.Empty;
            return value.Length <= maxChars ? value : value.Substring(0, maxChars);
        }
    }
}
