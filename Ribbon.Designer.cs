namespace PowerSyntax
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.languageDropDown = this.Factory.CreateRibbonDropDown();
            this.themeDropDown = this.Factory.CreateRibbonDropDown();
            this.ApplyButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.languageDropDown);
            this.group1.Items.Add(this.themeDropDown);
            this.group1.Items.Add(this.ApplyButton);
            this.group1.Label = "PowerSyntax";
            this.group1.Name = "group1";
            // 
            // languageDropDown
            // 
            this.languageDropDown.Label = "Language";
            this.languageDropDown.Name = "languageDropDown";
            this.languageDropDown.OfficeImageId = "FieldCodes";
            this.languageDropDown.ShowImage = true;
            this.languageDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LanguageDropDown_SelectionChanged);
            // 
            // themeDropDown
            // 
            this.themeDropDown.Label = "Theme";
            this.themeDropDown.Name = "themeDropDown";
            this.themeDropDown.OfficeImageId = "ThemeColorsGallery";
            this.themeDropDown.ShowImage = true;
            this.themeDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ThemeDropDown_SelectionChanged);
            // 
            // ApplyButton
            // 
            this.ApplyButton.Label = "Apply Syntax Highlight";
            this.ApplyButton.Name = "ApplyButton";
            this.ApplyButton.OfficeImageId = "StyleGalleryClassic";
            this.ApplyButton.ShowImage = true;
            this.ApplyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ApplyButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ApplyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown languageDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown themeDropDown;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
