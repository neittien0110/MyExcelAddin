namespace ExcelAddIn
{
    /// <summary>
    ///       Class thuộc kiểu Ribbon (Visual Designer)
    /// </summary>
    partial class MyRibon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyRibon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonImage2Cells = this.Factory.CreateRibbonButton();
            this.groupAlgorithm = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.editSaturationPeak = this.Factory.CreateRibbonEditBox();
            this.groupAudio = this.Factory.CreateRibbonGroup();
            this.buttonTextToSpeech = this.Factory.CreateRibbonButton();
            this.buttonSpeechVN = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupAlgorithm.SuspendLayout();
            this.groupAudio.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.groupAlgorithm);
            this.tab1.Groups.Add(this.groupAudio);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonImage2Cells);
            this.group1.Label = "Nghệ thuật";
            this.group1.Name = "group1";
            // 
            // buttonImage2Cells
            // 
            this.buttonImage2Cells.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonImage2Cells.Description = "Haha";
            this.buttonImage2Cells.Label = "Tô màu cell";
            this.buttonImage2Cells.Name = "buttonImage2Cells";
            this.buttonImage2Cells.OfficeImageId = "AllCategories";
            this.buttonImage2Cells.ScreenTip = "Chuyển ảnh thành cell";
            this.buttonImage2Cells.ShowImage = true;
            this.buttonImage2Cells.SuperTip = "Mỗi pixcel ảnh sẽ trở thành một cell trên excel. Ảnh được tự động co sao cho số đ" +
    "iểm ảnh không quá 82455 do giới hạn của Excel";
            this.buttonImage2Cells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImage2Cells_Click);
            // 
            // groupAlgorithm
            // 
            this.groupAlgorithm.Items.Add(this.button1);
            this.groupAlgorithm.Items.Add(this.dropDown1);
            this.groupAlgorithm.Items.Add(this.editSaturationPeak);
            this.groupAlgorithm.Label = "Thuật toán";
            this.groupAlgorithm.Name = "groupAlgorithm";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "Màu hóa";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "BlackAndWhiteLightGrayscale";
            this.button1.ScreenTip = "Màu hóa ma trận giá trị";
            this.button1.ShowImage = true;
            this.button1.SuperTip = "Lựa chọn một bảng, sau đó bấm nút Màu hóa. Các cell trong bảng sẽ được tô màu với" +
    " mức xám thay đổi  từ màu đen (0) tới mức cực đại, của màu chỉ định trong dropbo" +
    "x";
            // 
            // dropDown1
            // 
            ribbonDropDownItemImpl1.Label = "Đỏ";
            ribbonDropDownItemImpl1.OfficeImageId = "AppointmentColor1";
            ribbonDropDownItemImpl2.Label = "Xanh lá";
            ribbonDropDownItemImpl2.OfficeImageId = "AppointmentColor3";
            ribbonDropDownItemImpl3.Label = "Xanh dương";
            ribbonDropDownItemImpl3.OfficeImageId = "AppointmentColor2";
            this.dropDown1.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown1.Label = "Màu";
            this.dropDown1.Name = "dropDown1";
            // 
            // editSaturationPeak
            // 
            this.editSaturationPeak.Label = "Cực đại";
            this.editSaturationPeak.Name = "editSaturationPeak";
            this.editSaturationPeak.Text = "255";
            // 
            // groupAudio
            // 
            ribbonDialogLauncherImpl1.OfficeImageId = "FontsReplaceFonts";
            this.groupAudio.DialogLauncher = ribbonDialogLauncherImpl1;
            this.groupAudio.Items.Add(this.buttonTextToSpeech);
            this.groupAudio.Items.Add(this.buttonSpeechVN);
            this.groupAudio.Label = "Âm thanh/Giọng nói";
            this.groupAudio.Name = "groupAudio";
            this.groupAudio.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.groupAudio_DialogLauncherClick);
            // 
            // buttonTextToSpeech
            // 
            this.buttonTextToSpeech.Image = ((System.Drawing.Image)(resources.GetObject("buttonTextToSpeech.Image")));
            this.buttonTextToSpeech.Label = "Đọc tiếng Anh";
            this.buttonTextToSpeech.Name = "buttonTextToSpeech";
            this.buttonTextToSpeech.OfficeImageId = "SoundInsertFromFile";
            this.buttonTextToSpeech.ScreenTip = "Đọc văn bản đã chọn bằng tiếng Anh";
            this.buttonTextToSpeech.ShowImage = true;
            this.buttonTextToSpeech.SuperTip = "Chọn một hoặc nhiều cell cần đọc, và bấm nút này. Hàm sẽ kết nối với dịch vụ Cont" +
    "ana và không cần kết nối internet";
            this.buttonTextToSpeech.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTextToSpeech_Click);
            // 
            // buttonSpeechVN
            // 
            this.buttonSpeechVN.Image = ((System.Drawing.Image)(resources.GetObject("buttonSpeechVN.Image")));
            this.buttonSpeechVN.Label = "Đọc tiếng Việt";
            this.buttonSpeechVN.Name = "buttonSpeechVN";
            this.buttonSpeechVN.OfficeImageId = "SoundInsertFromFile";
            this.buttonSpeechVN.ScreenTip = "Đọc văn bản đã chọn bằng tiếng Việt";
            this.buttonSpeechVN.ShowImage = true;
            this.buttonSpeechVN.SuperTip = "Chọn một hoặc nhiều cell, rồi bấm nút này để nghe giọng nói tiếng Việt. Hàm sử dụ" +
    "ng dịch vụ tổng hợp tiếng nói tại website code.responsivevoice.org  nên nhất thi" +
    "ết phải có kết nối internet";
            this.buttonSpeechVN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSpeechVN_Click);
            // 
            // MyRibon
            // 
            this.Name = "MyRibon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupAlgorithm.ResumeLayout(false);
            this.groupAlgorithm.PerformLayout();
            this.groupAudio.ResumeLayout(false);
            this.groupAudio.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonImage2Cells;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAlgorithm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editSaturationPeak;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAudio;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTextToSpeech;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSpeechVN;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibon MyRibon
        {
            get { return this.GetRibbon<MyRibon>(); }
        }
    }
}
