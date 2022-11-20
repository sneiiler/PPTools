namespace PPTools
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.line_spacing_12 = this.Factory.CreateRibbonButton();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.line_spacing_specific = this.Factory.CreateRibbonButton();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.font_weiruanyahei = this.Factory.CreateRibbonButton();
            this.font_timesNR = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.delete_current_page_animation = this.Factory.CreateRibbonButton();
            this.delete_selected_page_animation = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.isLockAspectRatio = this.Factory.CreateRibbonCheckBox();
            this.resize_shape_fullbackground = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "PPTools";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.line_spacing_12);
            this.group1.Items.Add(this.editBox1);
            this.group1.Items.Add(this.line_spacing_specific);
            this.group1.Items.Add(this.buttonGroup1);
            this.group1.Label = "文本调整";
            this.group1.Name = "group1";
            // 
            // line_spacing_12
            // 
            this.line_spacing_12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.line_spacing_12.Image = ((System.Drawing.Image)(resources.GetObject("line_spacing_12.Image")));
            this.line_spacing_12.Label = "行距1.2";
            this.line_spacing_12.Name = "line_spacing_12";
            this.line_spacing_12.ScreenTip = "将行距调整为1.2倍";
            this.line_spacing_12.ShowImage = true;
            this.line_spacing_12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.line_spacing_12_Click);
            // 
            // editBox1
            // 
            this.editBox1.Image = ((System.Drawing.Image)(resources.GetObject("editBox1.Image")));
            this.editBox1.Label = "自定义行距";
            this.editBox1.Name = "editBox1";
            this.editBox1.ShowImage = true;
            this.editBox1.Text = "1.2";
            this.editBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox1_TextChanged);
            // 
            // line_spacing_specific
            // 
            this.line_spacing_specific.Image = ((System.Drawing.Image)(resources.GetObject("line_spacing_specific.Image")));
            this.line_spacing_specific.Label = "指定行间距";
            this.line_spacing_specific.Name = "line_spacing_specific";
            this.line_spacing_specific.ShowImage = true;
            this.line_spacing_specific.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.line_spacing_specific_Click);
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.font_weiruanyahei);
            this.buttonGroup1.Items.Add(this.font_timesNR);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // font_weiruanyahei
            // 
            this.font_weiruanyahei.Image = ((System.Drawing.Image)(resources.GetObject("font_weiruanyahei.Image")));
            this.font_weiruanyahei.Label = "微软雅黑";
            this.font_weiruanyahei.Name = "font_weiruanyahei";
            this.font_weiruanyahei.ShowImage = true;
            this.font_weiruanyahei.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.font_weiruanyahei_Click);
            // 
            // font_timesNR
            // 
            this.font_timesNR.Image = ((System.Drawing.Image)(resources.GetObject("font_timesNR.Image")));
            this.font_timesNR.Label = "Times NR";
            this.font_timesNR.Name = "font_timesNR";
            this.font_timesNR.ShowImage = true;
            this.font_timesNR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.font_timesNR_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.delete_current_page_animation);
            this.group2.Items.Add(this.delete_selected_page_animation);
            this.group2.Label = "动画相关";
            this.group2.Name = "group2";
            // 
            // delete_current_page_animation
            // 
            this.delete_current_page_animation.Image = ((System.Drawing.Image)(resources.GetObject("delete_current_page_animation.Image")));
            this.delete_current_page_animation.Label = "一键删除本页动画";
            this.delete_current_page_animation.Name = "delete_current_page_animation";
            this.delete_current_page_animation.ShowImage = true;
            this.delete_current_page_animation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.delete_current_page_animation_Click);
            // 
            // delete_selected_page_animation
            // 
            this.delete_selected_page_animation.Image = ((System.Drawing.Image)(resources.GetObject("delete_selected_page_animation.Image")));
            this.delete_selected_page_animation.Label = "删除选中页面动画";
            this.delete_selected_page_animation.Name = "delete_selected_page_animation";
            this.delete_selected_page_animation.ShowImage = true;
            this.delete_selected_page_animation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.delete_selected_page_animation_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.isLockAspectRatio);
            this.group3.Items.Add(this.resize_shape_fullbackground);
            this.group3.Label = "形状";
            this.group3.Name = "group3";
            // 
            // isLockAspectRatio
            // 
            this.isLockAspectRatio.Label = "保持原比例";
            this.isLockAspectRatio.Name = "isLockAspectRatio";
            // 
            // resize_shape_fullbackground
            // 
            this.resize_shape_fullbackground.Image = ((System.Drawing.Image)(resources.GetObject("resize_shape_fullbackground.Image")));
            this.resize_shape_fullbackground.Label = "尺寸最大并居中";
            this.resize_shape_fullbackground.Name = "resize_shape_fullbackground";
            this.resize_shape_fullbackground.ShowImage = true;
            this.resize_shape_fullbackground.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.resize_shape_fullbackground_Click);
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
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton line_spacing_12;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton line_spacing_specific;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton delete_current_page_animation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton delete_selected_page_animation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton font_weiruanyahei;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton font_timesNR;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resize_shape_fullbackground;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox isLockAspectRatio;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
