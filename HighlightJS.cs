using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Scripting.JavaScript;
using System.IO;

namespace PowerSyntax
{
    public class HighlightJS : IDisposable
    {
        public HighlightJS(IResourceProvider resource)
        {
            this.resource = resource;

            runtime = new JavaScriptRuntime();
            engine = runtime.CreateEngine();
            context = engine.AcquireContext();

            engine.SetGlobalFunction("require", Require);

            hljs = Require(engine, true, engine.UndefinedValue, Params(engine.Converter.FromString(@".\index"))) as JavaScriptObject;

            var listLanguages = hljs.GetPropertyByName("listLanguages") as JavaScriptFunction;
            languages = (listLanguages.Invoke(Params()) as JavaScriptArray).Select(x => engine.Converter.ToString(x)).OrderBy(x => x).ToArray();

            highlightJsFunction = engine.EvaluateScriptText(@"((hljs, code, languageSubset) => hljs.highlightAuto(code, languageSubset).value)").Invoke(Params()) as JavaScriptFunction;
        }

        public void Dispose()
        {
            highlightJsFunction.Dispose();
            hljs.Dispose();
            context.Dispose();
            engine.Dispose();
            runtime.Dispose();
        }

        public IList<string> ListLanguages() => languages;

        public string HighlightAuto(string code, string language)
        {
            return string.IsNullOrEmpty(code) ? string.Empty : engine.Converter.ToString(HighlightAuto(engine.Converter.FromString(code), engine.Converter.FromString(language)));
        }

        private JavaScriptValue HighlightAuto(JavaScriptValue code, JavaScriptValue language)
        {
            return highlightJsFunction.Invoke(Params(hljs, code, Array(language)));
        }

        private IEnumerable<JavaScriptValue> Params() => Enumerable.Empty<JavaScriptValue>();
        private IEnumerable<JavaScriptValue> Params(params JavaScriptValue[] args) => args;
        private JavaScriptArray Array() => engine.CreateArray(0);
        private JavaScriptArray Array(params JavaScriptValue[] args)
        {
            var array = engine.CreateArray(args.Length);
            for (int i = 0; i < args.Length; ++i)
            {
                array[i] = args[i];
            }
            return array;
        }

        private JavaScriptValue Require(JavaScriptEngine engine, bool construct, JavaScriptValue thisValue, IEnumerable<JavaScriptValue> arguments)
        {
            var args = arguments
                .Select(x => engine.Converter.ToString(x))
                .Select(x => Path.Combine($@"\lib\{x.Substring(2)}.js"))
                .ToArray();
            string code = resource.ReadAllText(args[0]);

            var old_module = engine.GetGlobalVariable("module");
            var old_export = engine.GetGlobalVariable("exports");
            try
            {
                var module = engine.CreateObject();
                var exports = engine.CreateObject();
                module.SetPropertyByName("exports", exports);
                engine.SetGlobalVariable("exports", exports);
                engine.SetGlobalVariable("module", module);
                engine.Execute(new Microsoft.Scripting.ScriptSource(args[0], code));
                var m_exports = module.GetPropertyByName("exports");
                return m_exports == exports ? exports : m_exports;
            }
            finally
            {
                engine.SetGlobalVariable("module", old_export);
                engine.SetGlobalVariable("exports", old_module);
            }
        }

        private JavaScriptRuntime runtime = null;
        private JavaScriptEngine engine = null;
        private JavaScriptExecutionContext context = null;
        private JavaScriptFunction highlightJsFunction = null;
        private JavaScriptObject hljs = null;

        private string[] languages = null;
        private IResourceProvider resource = null;
    }
}
