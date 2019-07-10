using System.Collections.Generic;
using System.Linq;
using AngleSharp;
using AngleSharp.Dom;
using System.Drawing;
using System.Globalization;
using AngleSharp.Css.Dom;
using System;

namespace PowerSyntax
{
    public class HighlightJSHighlighter : ISyntaxHighlighter, IDisposable
    {
        public HighlightJSHighlighter(IResourceProvider resource)
        {
            this.resource = resource;

            themes = resource.GetFiles("/styles", "*.css").Select(x => x.Split('.').Reverse().Skip(1).First()).ToArray();

            highlightJS = new HighlightJS(resource);

            var config = Configuration.Default.WithCss();
            using (var context = BrowsingContext.New(config) as BrowsingContext)
            {
                context.OpenAsync(req => req.Content("<html></html>"));
            }
        }

        public void Dispose()
        {
            highlightJS.Dispose();
        }

        public IList<string> Language => highlightJS.ListLanguages();
        public IList<string> Themes => themes;

        #region BrowserEngine

        static StyleRange ExtractStyle(ICssStyleDeclaration style)
        {
            var mappedStyle = new StyleRange();

            string cssColor = null;
            try
            {
                cssColor = style?.GetPropertyValue("color");
                mappedStyle.Color = ParseColor(cssColor);
            }
            catch (NullReferenceException) { }

            try
            {
                string fontWeight = null;
                fontWeight = style?.GetPropertyValue("font-weight");
                mappedStyle.Bold = (!string.IsNullOrEmpty(fontWeight) && fontWeight == "bold");
            }
            catch (NullReferenceException) { }


            try
            {
                string fontStyle = null;
                fontStyle = style?.GetPropertyValue("font-style");
                mappedStyle.Italic = (!string.IsNullOrEmpty(fontStyle) && fontStyle == "italic");
            }
            catch (NullReferenceException) { }

            return mappedStyle;
        }

        private StyleRange Parse(int begin, IElement element, IWindow browsingWindow)
        {
            StyleRange style;
            try
            {
                style = ExtractStyle(browsingWindow.GetComputedStyle(element));
            }
            catch (NullReferenceException)
            {
                style = new StyleRange();
            }

            style.Begin = begin;
            style.Length = element.TextContent.Length;
            style.Text = element.TextContent;

            style.Children = Parse(begin, element.ChildNodes, browsingWindow);

            return style;
        }

        private IEnumerable<StyleRange> Parse(int begin, INodeList nodes, IWindow browsingWindow)
        {
            var children = new List<StyleRange>();
            int sourcePosition = begin;
            foreach (var node in nodes)
            {
                if (node is IElement element)
                {
                    children.Add(Parse(sourcePosition, element, browsingWindow));
                }
                sourcePosition += node.TextContent.Length;
            }
            return children;
        }

        public StyleRange Parse(string code, string language, string theme)
        {
            string highlighted = ApplyHighlight(code, language);
            string css = resource.ReadAllText($"/styles/{theme}.css");
            var source = $"<html><style type=\"text/css\">{css}</style></head><body><div class=\"hljs\">{highlighted}</div></body></html>";

            var config = Configuration.Default.WithCss();
            using (var context = BrowsingContext.New(config) as BrowsingContext)
            {
                var task = context.OpenAsync(req => req.Content(source));
                task.Wait();
                var document = task.Result;
                var hljsbody = document.Body.Children.First();
                var style = Parse(0, hljsbody, document.DefaultView);

                string cssColor = null;
                try
                {
                    var cssStyle = document.DefaultView.GetComputedStyle(hljsbody);
                    cssColor = cssStyle?.GetPropertyValue("background");
                }
                catch (NullReferenceException) { }
                style.BackGroundColor = ParseColor(cssColor);
                return style;
            }
        }

        private static Color? ParseColor(string cssColor)
        {
            if (string.IsNullOrEmpty(cssColor))
                return null;

            try
            {
                cssColor = cssColor.Trim();

                if (cssColor.Contains("rgb")) //rgb or argb
                {
                    int left = cssColor.IndexOf('(');
                    int right = cssColor.IndexOf(')');

                    if (left < 0 || right < 0)
                        throw null;

                    string noBrackets = cssColor.Substring(left + 1, right - left - 1);

                    string[] parts = noBrackets.Split(',');

                    int r = int.Parse(parts[0], CultureInfo.InvariantCulture);
                    int g = int.Parse(parts[1], CultureInfo.InvariantCulture);
                    int b = int.Parse(parts[2], CultureInfo.InvariantCulture);

                    if (parts.Length == 3)
                    {
                        return Color.FromArgb(r, g, b);
                    }
                    else if (parts.Length == 4)
                    {
                        float a = float.Parse(parts[3], CultureInfo.InvariantCulture);
                        return Color.FromArgb((int)(a * 255), r, g, b);
                    }
                    return null;
                }

                return ColorTranslator.FromHtml(cssColor);
            }
            catch (FormatException)
            {
            }
            return null;
        }

        #endregion

        private string ApplyHighlight(string code, string language)
        {
            return string.IsNullOrEmpty(code) ? string.Empty : highlightJS.HighlightAuto(code, language);
        }

        private HighlightJS highlightJS = null;

        private string[] themes = null;
        private IResourceProvider resource = null;
    }

}
