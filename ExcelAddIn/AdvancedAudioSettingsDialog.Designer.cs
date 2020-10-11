namespace ExcelAddIn
{
    partial class AdvancedAudioSettingsDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.propertyGridAudioSetting = new System.Windows.Forms.PropertyGrid();
            this.SuspendLayout();
            // 
            // propertyGridAudioSetting
            // 
            this.propertyGridAudioSetting.Location = new System.Drawing.Point(12, 12);
            this.propertyGridAudioSetting.Name = "propertyGridAudioSetting";
            this.propertyGridAudioSetting.Size = new System.Drawing.Size(776, 426);
            this.propertyGridAudioSetting.TabIndex = 0;
            // 
            // AdvancedAudioSettingsDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.propertyGridAudioSetting);
            this.Name = "AdvancedAudioSettingsDialog";
            this.Text = "AdvancedAudioSettingsDialog";
            this.Load += new System.EventHandler(this.AdvancedAudioSettingsDialog_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PropertyGrid propertyGridAudioSetting;
    }
}