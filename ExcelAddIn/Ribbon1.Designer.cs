namespace ExcelAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl2 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.systemGroup = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "BIM应用";
            this.tab1.Name = "tab1";
            this.tab1.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabInsert");
            // 
            // group
            // 
            this.group.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group.Items.Add(this.systemGroup);
            this.group.Items.Add(this.separator1);
            this.group.Items.Add(this.button1);
            this.group.Label = "数据库";
            this.group.Name = "group";
            this.group.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group1_DialogLauncherClick);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group1
            // 
            this.group1.DialogLauncher = ribbonDialogLauncherImpl2;
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.button3);
            this.group1.Label = "数据库";
            this.group1.Name = "group1";
            this.group1.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group1_DialogLauncherClick);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::ExcelAddIn.Properties.Resources._1145566;
            this.button2.Label = "运行参数";
            this.button2.Name = "button2";
            this.button2.OfficeImageId = "Call";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.System_Group_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = global::ExcelAddIn.Properties.Resources._1145595;
            this.button4.Label = "二级系统";
            this.button4.Name = "button4";
            this.button4.OfficeImageId = "Call";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.System_Group_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::ExcelAddIn.Properties.Resources._1170739;
            this.button3.Label = "三级系统";
            this.button3.Name = "button3";
            this.button3.OfficeImageId = "Call";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.System_Group_Click);
            // 
            // systemGroup
            // 
            this.systemGroup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.systemGroup.Image = global::ExcelAddIn.Properties.Resources._1145472;
            this.systemGroup.Label = "系统分组";
            this.systemGroup.Name = "systemGroup";
            this.systemGroup.OfficeImageId = "Call";
            this.systemGroup.ShowImage = true;
            this.systemGroup.SuperTip = "格式化文本为SQL需要的格式";
            this.systemGroup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.System_Group_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::ExcelAddIn.Properties.Resources._1145507;
            this.button1.Label = "系统分类";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "Call";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.System_Group_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group.ResumeLayout(false);
            this.group.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton systemGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
