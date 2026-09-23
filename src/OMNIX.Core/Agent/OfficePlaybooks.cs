using System;
using System.IO;
using System.Reflection;
using Newtonsoft.Json.Linq;
using OMNIX.Core.Context;
namespace OMNIX.Core.Agent
{
    public static class OfficePlaybooks
    {
        private static string Read(string name)
        {
            using(var stream=typeof(OfficePlaybooks).Assembly.GetManifestResourceStream("OMNIX.Agent."+name))
            {
                if(stream==null) throw new InvalidOperationException("Missing Office playbook: "+name);
                using(var reader=new StreamReader(stream)) return reader.ReadToEnd();
            }
        }
        public static string Load(HostType host,string request)
        {
            string result=Read("contract.md")+"\n"+Read(host+".md");
            string q=(request??"").ToLowerInvariant();
            foreach(string term in new[]{"shop","sale","invoice","stock","gold","inventory","business","دکان","طلا","فروش","مشتری","انبار","دیتابیس","database"})
                if(q.Contains(term)) { result+="\n"+Read("business.md"); break; }
            return result;
        }
        public static string Template(string name,string sheet,HostType host)
        {
            if(host!=HostType.Excel) throw new ArgumentException("Typed templates currently available for Excel: gold, inventory, invoice. Use the host guide for other documents.");
            if(name!="gold" && name!="inventory" && name!="invoice") throw new ArgumentException("Available templates: gold, inventory, invoice.");
            if(string.IsNullOrWhiteSpace(sheet)) throw new ArgumentException("An exact new sheet name is required.");
            var plan=JObject.Parse(Read(name+".json"));
            foreach(JObject step in (JArray)plan["steps"])
            {
                step["args"]["sheet"]=sheet;
                foreach(JObject check in (JArray)step["checks"]) check["sheet"]=sheet;
            }
            return plan.ToString(Newtonsoft.Json.Formatting.None);
        }
    }
}
