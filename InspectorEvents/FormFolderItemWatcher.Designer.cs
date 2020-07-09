namespace InspectorEvents
{
    partial class FormFolderItemWatcher
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
            this.listBoxFolderItems = new System.Windows.Forms.ListBox();
            this.checkBoxAlwaysOnTop = new System.Windows.Forms.CheckBox();
            this.buttonClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listBoxFolderItems
            // 
            this.listBoxFolderItems.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBoxFolderItems.FormattingEnabled = true;
            this.listBoxFolderItems.Location = new System.Drawing.Point(12, 12);
            this.listBoxFolderItems.Name = "listBoxFolderItems";
            this.listBoxFolderItems.Size = new System.Drawing.Size(495, 329);
            this.listBoxFolderItems.TabIndex = 0;
            // 
            // checkBoxAlwaysOnTop
            // 
            this.checkBoxAlwaysOnTop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBoxAlwaysOnTop.AutoSize = true;
            this.checkBoxAlwaysOnTop.Checked = true;
            this.checkBoxAlwaysOnTop.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxAlwaysOnTop.Location = new System.Drawing.Point(12, 349);
            this.checkBoxAlwaysOnTop.Name = "checkBoxAlwaysOnTop";
            this.checkBoxAlwaysOnTop.Size = new System.Drawing.Size(92, 17);
            this.checkBoxAlwaysOnTop.TabIndex = 4;
            this.checkBoxAlwaysOnTop.Text = "Always on top";
            this.checkBoxAlwaysOnTop.UseVisualStyleBackColor = true;
            this.checkBoxAlwaysOnTop.CheckedChanged += new System.EventHandler(this.checkBoxAlwaysOnTop_CheckedChanged);
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.Location = new System.Drawing.Point(432, 345);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(75, 23);
            this.buttonClose.TabIndex = 3;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // FormFolderItemWatcher
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(519, 376);
            this.Controls.Add(this.checkBoxAlwaysOnTop);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.listBoxFolderItems);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormFolderItemWatcher";
            this.Text = "Folder Item Watcher";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxFolderItems;
        private System.Windows.Forms.CheckBox checkBoxAlwaysOnTop;
        private System.Windows.Forms.Button buttonClose;
    }
}