using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OMNIX.Core.Context;
using OMNIX.Core.Tools;

namespace OMNIX.Core.Agent
{
    public interface IPlanVerificationHost
    {
        // Null means passed. Implementations inspect native Office state on its owner thread.
        string CheckPostcondition(JObject check);
    }

    public sealed class ExecutionPlan
    {
        private JObject _plan;
        private readonly HashSet<string> _applied = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _passed = new HashSet<string>(StringComparer.Ordinal);
        private readonly Dictionary<string,int> _attempts = new Dictionary<string,int>(StringComparer.Ordinal);
        public string OriginalRequest { get; private set; }
        public bool Required { get; private set; }
        public bool Complete { get { return _plan != null && Steps.All(s => _passed.Contains((string)s["id"])); } }
        public string PreviousCheckpoint { get; set; }
        public Action<string> SaveCheckpoint { get; set; }
        private IEnumerable<JObject> Steps { get { return _plan == null ? Enumerable.Empty<JObject>() : ((JArray)_plan["steps"]).Cast<JObject>(); } }

        public void Begin(string request, bool required)
        {
            OriginalRequest = request ?? ""; Required = required; _plan = null;
            _applied.Clear(); _passed.Clear(); _attempts.Clear();
        }

        public string Submit(string json, HostType host)
        {
            if (json == null || json.Length > 60000) throw new ArgumentException("Plan exceeds 60,000 characters.");
            var plan = JObject.Parse(json);
            var steps = plan["steps"] as JArray;
            if (steps == null || steps.Count < 1 || steps.Count > 12) throw new ArgumentException("Plan requires 1–12 ordered steps.");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var token in steps)
            {
                var step = token as JObject;
                if (step == null) throw new ArgumentException("Each step must be an object.");
                string id = (string)step["id"], tool = (string)step["tool"];
                if (string.IsNullOrWhiteSpace(id) || id.Length > 40 || !ids.Add(id)) throw new ArgumentException("Step IDs must be unique, nonempty and at most 40 characters.");
                if (!ToolNames.IsWriteTool(tool) || !(step["args"] is JObject)) throw new ArgumentException("Step requires a real write tool and args object.");
                var checks = step["checks"] as JArray;
                if (checks == null || checks.Count < 1 || checks.Count > 12) throw new ArgumentException("Each step requires 1–12 native postconditions.");
                foreach (var c in checks) ValidateCheck(c as JObject, host);
                if (tool == ToolNames.CreateDataTable && (bool?)step["args"]["uniqueName"] == true)
                    throw new ArgumentException("Planned creation requires an exact, unused sheet name; automatic suffixes are not allowed.");
            }
            // A model cannot erase already-executed goals or weaken checks to declare success.
            if (_plan != null && _applied.Count > 0)
            {
                foreach (var old in Steps)
                {
                    var next = steps.OfType<JObject>().FirstOrDefault(s => (string)s["id"] == (string)old["id"]);
                    if (next == null || !JToken.DeepEquals(old["checks"], next["checks"]))
                        throw new ArgumentException("After execution starts, retain every step ID and its original postconditions. Repair the operation, not the acceptance criteria.");
                    if (_passed.Contains((string)old["id"]) && !JToken.DeepEquals(old, next))
                        throw new ArgumentException("Do not change an already verified step.");
                }
            }
            _plan = plan;
            Checkpoint();
            return "Plan accepted. Execute the next pending step exactly; a native check follows each write. " + Summary();
        }

        public static void ValidateCheck(JObject c, HostType host)
        {
            if (c == null) throw new ArgumentException("Postcondition must be an object.");
            string kind = (string)c["kind"];
            string[] allowed = host == HostType.Excel ? new[]{"cell_value","formula","heading","table","no_errors"} :
                host == HostType.Word ? new[]{"text","paragraph_style","table_count"} : new[]{"text","slide_count","shape_bounds"};
            if (!allowed.Contains(kind)) throw new ArgumentException("Unsupported postcondition for " + host + ": " + kind);
            if (host == HostType.Excel && (string.IsNullOrWhiteSpace((string)c["sheet"]) || string.IsNullOrWhiteSpace((string)c["address"])))
                throw new ArgumentException("Excel checks require exact sheet and address.");
            if ((kind == "cell_value" || kind == "formula") && c["value"] == null)
                throw new ArgumentException("Cell/formula checks require the expected computed value.");
            if ((kind == "heading" || kind == "text") && string.IsNullOrWhiteSpace((string)c["text"]))
                throw new ArgumentException("Text checks require expected nonempty text.");
            if (kind == "paragraph_style" && string.IsNullOrWhiteSpace((string)c["style"]))
                throw new ArgumentException("Style check requires the expected style name.");
            if (kind == "formula" && string.IsNullOrWhiteSpace((string)c["formula"]))
                throw new ArgumentException("Formula checks require the exact formula and expected result.");
            if (kind == "table_count" || kind == "slide_count")
                if (c["count"] == null || c["count"].Type != JTokenType.Integer || (int)c["count"] < 0)
                    throw new ArgumentException("Count checks require an explicit nonnegative integer.");
        }

        public string BeforeWrite(ToolCall call)
        {
            if (!Required) return null;
            if (_plan == null) return "PLAN REQUIRED: submit_execution_plan before any write. Include exact tool arguments and native postconditions.";
            var step = Steps.FirstOrDefault(s => !_passed.Contains((string)s["id"]));
            if (step == null) return "Plan already completed. Do not repeat writes.";
            JObject args;
            try { args = JObject.Parse(call.ArgumentsJson ?? "{}"); } catch { return "Invalid write arguments."; }
            if ((string)step["tool"] != call.Name || !JToken.DeepEquals(step["args"],args))
                return "Write differs from the next pending plan step. Inspect the current state, then revise the pending operation if needed. " + Summary();
            string id = (string)step["id"];
            int attempts; _attempts.TryGetValue(id,out attempts);
            if (attempts >= 3) return "Step exhausted its three execution attempts. Stop and report incomplete; do not create a duplicate deliverable.";
            // Replaying an additive write after successful apply but failed verification duplicates content.
            if (_applied.Contains(id) && call.Name == ToolNames.CreateDataTable)
                return "This creation already applied. Revise this pending step to repair the existing sheet, retaining its checks.";
            _attempts[id] = attempts + 1;
            return null;
        }

        public string AfterWrite(IPlanVerificationHost host)
        {
            if (!Required || _plan == null) return "";
            var step = Steps.FirstOrDefault(s => !_passed.Contains((string)s["id"]));
            if (step == null) return Summary();
            string id = (string)step["id"]; _applied.Add(id);
            var failures = Check(host, step);
            if (failures.Count == 0) _passed.Add(id);
            VerifyAll(host); // A later write can invalidate an earlier accepted result.
            return failures.Count == 0 ? "POSTCONDITIONS PASSED: " + id + ". " + Summary() :
                "CHANGE APPLIED, POSTCONDITIONS FAILED: " + id + ". Repair existing content. " + string.Join("; ", failures);
        }

        public string VerifyAll(IPlanVerificationHost host)
        {
            if (_plan == null) return "No execution plan exists.";
            foreach (var step in Steps)
            {
                string id = (string)step["id"];
                if (!_applied.Contains(id)) continue;
                if (Check(host,step).Count == 0) _passed.Add(id); else _passed.Remove(id);
            }
            Checkpoint(); return Summary();
        }
        private static List<string> Check(IPlanVerificationHost host,JObject step)
        {
            var failures=new List<string>();
            foreach(var token in (JArray)step["checks"])
            {
                try { string failure=host.CheckPostcondition((JObject)token); if(failure!=null) failures.Add(failure); }
                catch(Exception) { failures.Add("Native check could not complete for " + (string)token["kind"] + "; inspect the target."); }
            }
            return failures;
        }
        public string Summary()
        {
            return _plan == null ? "Plan: required before writing." : "Plan status: " + (Complete ? "verified" : _applied.Count>0 ? "partial" : "planned") + "; " +
                string.Join(", ",Steps.Select(s=>(string)s["id"]+":"+(_passed.Contains((string)s["id"])?"verified":_applied.Contains((string)s["id"])?"needs_repair":"pending")));
        }
        public string Snapshot()
        {
            return new JObject { ["request"]=OriginalRequest, ["plan"]=_plan, ["status"]=Summary() }.ToString(Formatting.None);
        }
        private void Checkpoint()
        {
            // Persistence failure must not turn an applied Office operation into a retryable write failure.
            try { if(SaveCheckpoint!=null) SaveCheckpoint(Snapshot()); }
            catch(Exception) { /* Native verification remains authoritative. */ }
        }
        public string Envelope()
        {
            return "EXECUTION CONTRACT: " + Summary() + "\nOriginal user request: " + OriginalRequest +
                (_plan==null ? "" : "\nActive immutable acceptance plan: " + _plan.ToString(Formatting.None)) +
                (string.IsNullOrEmpty(PreviousCheckpoint)?"":"\nPrevious task checkpoint (context only; re-inspect before resuming): " + PreviousCheckpoint);
        }
    }
}
