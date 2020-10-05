using System.Linq;

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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupImportImage = this.Factory.CreateRibbonGroup();
            this.buttonImage2Cells = this.Factory.CreateRibbonButton();
            this.groupAlgorithm = this.Factory.CreateRibbonGroup();
            this.buttonColorize = this.Factory.CreateRibbonButton();
            this.dropDownColorRGB = this.Factory.CreateRibbonDropDown();
            this.editSaturationPeak = this.Factory.CreateRibbonEditBox();
            //this.groupFontBlackWhite = this.Factory.CreateRibbonGroup();
            //this.buttonFontBlackWhite = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupImportImage.SuspendLayout();
            this.groupAlgorithm.SuspendLayout();
            //this.groupFontBlackWhite.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupImportImage);
            this.tab1.Groups.Add(this.groupAlgorithm);
            //this.tab1.Groups.Add(this.groupFontBlackWhite);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupImportImage
            // 
            this.groupImportImage.Items.Add(this.buttonImage2Cells);
            this.groupImportImage.Label = "Import Image";
            this.groupImportImage.Name = "groupImportImage";
            // 
            // buttonImage2Cells
            // 
            this.buttonImage2Cells.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonImage2Cells.Description = "Haha";
            this.buttonImage2Cells.Label = "Load Image";
            this.buttonImage2Cells.Name = "buttonImage2Cells";
            this.buttonImage2Cells.OfficeImageId = "AllCategories";
            this.buttonImage2Cells.ScreenTip = "Chuyển ảnh thành cell";
            this.buttonImage2Cells.ShowImage = true;
            this.buttonImage2Cells.SuperTip = "Mỗi pixcel ảnh sẽ trở thành một cell trên excel. Ảnh được tự động co sao cho số đ" +
    "iểm ảnh không quá 82455 do giới hạn của Excel";
            this.buttonImage2Cells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImage2Cells_Click);

            //
            // groupFontBlackWhite
            //
            //this.groupFontBlackWhite.Items.Add(this.buttonFontBlackWhite);
            //this.groupFontBlackWhite.Label = "Black White Font";
            //this.groupFontBlackWhite.Name = "groupFontBlackWhite";
            //
            // buttonFontBlackWhite
            //
            //this.buttonFontBlackWhite.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            //this.buttonFontBlackWhite.Description = "???";
            //this.buttonFontBlackWhite.Label = "Change";
            //this.buttonFontBlackWhite.Name = "buttonFontBlackWhite";
            //this.buttonFontBlackWhite.OfficeImageId = "AllCategories";
            //this.buttonFontBlackWhite.ScreenTip = "Chuyển màu font";
            //this.buttonFontBlackWhite.ShowImage = true;
            //this.buttonFontBlackWhite.SuperTip = "Chuyển màu font chữ sang trắng hoặc đen";
            //this.buttonFontBlackWhite.Click += (sender, e) => this.buttonFontBlackWhite_Click();

            // 
            // groupAlgorithm
            // 
            this.groupAlgorithm.Items.Add(this.buttonColorize);
            this.groupAlgorithm.Items.Add(this.dropDownColorRGB);
            this.groupAlgorithm.Items.Add(this.editSaturationPeak);
            this.groupAlgorithm.Label = "Algorithm";
            this.groupAlgorithm.Name = "groupAlgorithm";
            // 
            // buttonCorlorize
            // 
            this.buttonColorize.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonColorize.Label = "Màu hóa";
            this.buttonColorize.Name = "button1";
            this.buttonColorize.OfficeImageId = "BlackAndWhiteLightGrayscale";
            this.buttonColorize.ScreenTip = "Màu hóa ma trận giá trị";
            this.buttonColorize.ShowImage = true;
            this.buttonColorize.SuperTip = "Lựa chọn một bảng, sau đó bấm nút Màu hóa. Các cell trong bảng sẽ được tô màu với" +
    " mức xám thay đổi  từ màu đen (0) tới mức cực đại, của màu chỉ định trong dropbo" +
    "x";

            //this code cannot pass arguments
            //this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonColorize);

            //use this code to pass arguments: color and saturation
            this.buttonColorize.Click += (sender,e) => this.buttonColorize_Click(dropDownColorRGB.SelectedItem.OfficeImageId,editSaturationPeak.Text);

            // 
            // dropDown1
            // 
            ribbonDropDownItemImpl1.Label = "Đỏ";
            ribbonDropDownItemImpl1.OfficeImageId = "AppointmentColor1";
            ribbonDropDownItemImpl2.Label = "Xanh lá";
            ribbonDropDownItemImpl2.OfficeImageId = "AppointmentColor3";
            ribbonDropDownItemImpl3.Label = "Xanh dương";
            ribbonDropDownItemImpl3.OfficeImageId = "AppointmentColor2";
            ribbonDropDownItemImpl4.Label = "Xám";
            ribbonDropDownItemImpl4.OfficeImageId = "AppointmentColor4";
            this.dropDownColorRGB.Items.Add(ribbonDropDownItemImpl1);
            this.dropDownColorRGB.Items.Add(ribbonDropDownItemImpl2);
            this.dropDownColorRGB.Items.Add(ribbonDropDownItemImpl3);
            this.dropDownColorRGB.Items.Add(ribbonDropDownItemImpl4);
            this.dropDownColorRGB.Label = "Màu";
            this.dropDownColorRGB.Name = "dropDown1";
            // 
            // editSaturationPeak
            // 
            this.editSaturationPeak.Label = "Cực đại";
            this.editSaturationPeak.Name = "editSaturationPeak";
            this.editSaturationPeak.Text = "255";
            // 
            // MyRibon
            // 
            this.Name = "MyRibon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupImportImage.ResumeLayout(false);
            this.groupImportImage.PerformLayout();
            this.groupAlgorithm.ResumeLayout(false);
            this.groupAlgorithm.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupImportImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonImage2Cells;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAlgorithm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonColorize;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownColorRGB;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editSaturationPeak;
        //internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFontBlackWhite;
        //internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFontBlackWhite;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibon MyRibon
        {
            get { return this.GetRibbon<MyRibon>(); }
        }
    }
}
