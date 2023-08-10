namespace RevitFamilyInstanceLock
{
    partial class FamilyPicker
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
            this.familySymbolListView = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // familySymbolListView
            // 
            this.familySymbolListView.HideSelection = false;
            this.familySymbolListView.Location = new System.Drawing.Point(17, 18);
            this.familySymbolListView.Margin = new System.Windows.Forms.Padding(4);
            this.familySymbolListView.Name = "familySymbolListView";
            this.familySymbolListView.Size = new System.Drawing.Size(634, 372);
            this.familySymbolListView.TabIndex = 0;
            this.familySymbolListView.UseCompatibleStateImageBehavior = false;
            // 
            // FamilyPicker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 638);
            this.Controls.Add(this.familySymbolListView);
            this.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FamilyPicker";
            this.Text = "FamilyConversionToDirectShape";
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.ListView familySymbolListView;
    }
}