using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Threading.Tasks;
using System.Drawing;

namespace PowerSyntax
{

    public static class Extension
    {
        public static Office.MsoTriState ToTryState(this bool state) => state ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
        public static bool ToBool(this Office.MsoTriState state) => state == Office.MsoTriState.msoTrue;
        public static int ToOle(this Color color) => ColorTranslator.ToOle(color);
    }

    public partial class ThisAddIn
    {
        private HighlightJSHighlighter Highlighter = null;
        private Ribbon Ribbon = null;

        private string SelectedLanguage => Ribbon.languageDropDown.SelectedItem.Label;
        private string SelectedTheme => Ribbon.themeDropDown.SelectedItem.Label;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Directory.GetCurrentDirectory();
            Highlighter = new HighlightJSHighlighter(new ManifestResourceProvider(@"PowerSyntax.highlight.node_modules.highlight.js"));
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Highlighter.Dispose();
        }

        private static void LoadDropDownItems(RibbonFactory factory, RibbonDropDown dropDown, IEnumerable<string> itemLabels, string savedItemLabel = null)
        {
            RibbonDropDownItem savedItem = null;

            foreach (var label in itemLabels)
            {
                var item = factory.CreateRibbonDropDownItem();
                item.Label = label;
                if (label == savedItemLabel)
                {
                    savedItem = item;
                }
                dropDown.Items.Add(item);
            }
            if (savedItem != null)
            {
                dropDown.SelectedItem = savedItem;
            }
        }

        internal void OnRibbonLoaded(Ribbon ribbon)
        {
            Ribbon = ribbon;
            LoadDropDownItems(Ribbon.Factory, Ribbon.languageDropDown, Highlighter.Language, Properties.Settings.Default.Language);
            LoadDropDownItems(Ribbon.Factory, Ribbon.themeDropDown, Highlighter.Themes, Properties.Settings.Default.Theme);
        }

        private static void ApplyStyle(PowerPoint.TextRange selection, StyleRange style)
        {
            var range = selection.Characters(style.Begin + 1, style.Length);
            if (style.Color is Color color)
            {
                range.Font.Color.RGB = color.ToOle();
            }
            if (style.Bold is bool bold)
            {
                range.Font.Bold = bold.ToTryState();
            }
            if (style.Italic is bool italic)
            {
                range.Font.Italic = italic.ToTryState();
            }

            ApplyStyle(selection, style.Children);
        }

        private static void ApplyStyle(PowerPoint.TextRange selection, IEnumerable<StyleRange> styles)
        {
            foreach (var style in styles)
            {
                ApplyStyle(selection, style);
            }
        }

        private void Normalize(PowerPoint.TextRange text)
        {
            for (int i = 1; i <= text.Length; ++i)
            {
                var c = text.Characters(i);
                switch (c.Text[0])
                {
                    case '“': c.Text = "\""; break;
                    case '”': c.Text = "\""; break;
                    case '‘': c.Text = "'"; break;
                    case '’': c.Text = "'"; break;
                    default: break;
                }
            }
        }

        private StyleRange Highlight(PowerPoint.TextRange text)
        {
            Normalize(text);
            var style = Highlighter.Parse(text.Text, SelectedLanguage, SelectedTheme);
            ApplyStyle(text, style);
            return style;
        }

        private void Highlight(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame.ToBool() && shape.TextFrame.HasText.ToBool())
            {
                var style = Highlight(shape.TextFrame.TextRange);
                if (style.BackGroundColor is Color color)
                {
                    shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
                }
            }
        }

        private void Highlight(IEnumerable<PowerPoint.Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    Highlight(shape.GroupItems.Cast<PowerPoint.Shape>());
                }
                else
                {
                    Highlight(shape);
                }
            }
        }

        private void Highlight(PowerPoint.SlideRange slides)
        {
            foreach (PowerPoint.Slide slide in slides)
            {
                Highlight(slide.Shapes.Cast<PowerPoint.Shape>());
            }
        }

        private void Highlight(PowerPoint.Selection selection)
        {
            switch (selection.Type)
            {
                default:
                case PowerPoint.PpSelectionType.ppSelectionNone:
                    break;
                case PowerPoint.PpSelectionType.ppSelectionSlides:
                    Highlight(selection.SlideRange);
                    break;
                case PowerPoint.PpSelectionType.ppSelectionShapes:
                    Highlight(selection.ShapeRange.Cast<PowerPoint.Shape>());
                    break;
                case PowerPoint.PpSelectionType.ppSelectionText:
                    if (selection.TextRange.Length > 0)
                    {
                        Highlight(selection.TextRange);
                    }
                    else
                    {
                        Highlight(selection.ShapeRange.Cast<PowerPoint.Shape>());
                    }
                    break;
            }
        }

        PowerPoint.Selection Selection {
            get
            {
                try
                {
                    return Globals.ThisAddIn.Application.ActiveWindow.Selection;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    return null;
                }
            }
        }

        public void ApplySyntaxHighlight()
        {
            var selection = Selection;
            if (selection != null)
            {
                Highlight(selection);
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
