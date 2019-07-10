using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace PowerSyntax
{
    public partial class Ribbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.OnRibbonLoaded(this);
        }

        private void ApplyButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ApplySyntaxHighlight();
        }

        private void LanguageDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (!string.IsNullOrEmpty(languageDropDown.SelectedItem.Label))
            {
                Properties.Settings.Default.Language = languageDropDown.SelectedItem.Label;
                Properties.Settings.Default.Save();
            }
        }

        private void ThemeDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (!string.IsNullOrEmpty(themeDropDown.SelectedItem.Label))
            {
                Properties.Settings.Default.Theme = themeDropDown.SelectedItem.Label;
                Properties.Settings.Default.Save();
            }
        }
    }
}
