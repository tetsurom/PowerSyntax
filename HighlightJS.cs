using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Microsoft.ClearScript.V8;
using AngleSharp.Dom;
using Microsoft.ClearScript;

namespace PowerSyntax
{
    public class HighlightJS : IDisposable
    {
        public HighlightJS(IResourceProvider resource)
        {
            this.resource = resource;

            engine = new V8ScriptEngine();

            engine.AddHostObject("require", new Func<string, object>(x => {
                string code = resource.ReadAllText(Path.Combine($@"\lib\{x.Substring(2)}.js"));
                return engine.Evaluate(new DocumentInfo(x), @"(function(){ var exports = {}; var module = { exports: exports };" + code + @"; return exports === module.exports ? exports : module.exports; })()");
            }));

            hljs = engine.Script.require(@".\index");

            var listLanguages = hljs.GetProperty("listLanguages") as ScriptObject;
            var languagesArray = listLanguages.Invoke(false) as ScriptObject;
            var length = languagesArray.GetProperty("length") as int? ?? 0;
            languages = Enumerable.Range(0, length).Select(x => languagesArray.GetProperty(x).ToString()).ToArray();

            highlightJsFunction = engine.Evaluate(@"((hljs, code, languageSubset) => hljs.highlightAuto(code, [languageSubset]).value)") as ScriptObject;
        }

        public void Dispose()
        {
            engine.Dispose();
        }

        public IList<string> ListLanguages() => languages;

        public string HighlightAuto(string code, string language)
        {
            return highlightJsFunction.Invoke(false, hljs, code, language).ToString();
        }

        private V8ScriptEngine engine = null;
        private ScriptObject highlightJsFunction = null;
        private ScriptObject hljs = null;

        private string[] languages = null;
        private IResourceProvider resource = null;
    }
}
