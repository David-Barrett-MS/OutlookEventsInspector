namespace InspectorEvents
{
    partial class FormEventTracker
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
            this.listBoxLog = new System.Windows.Forms.ListBox();
            this.buttonClose = new System.Windows.Forms.Button();
            this.checkBoxAlwaysOnTop = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBoxPropId = new System.Windows.Forms.TextBox();
            this.checkBoxRetrievePropertyonSelectionChange = new System.Windows.Forms.CheckBox();
            this.buttonClear = new System.Windows.Forms.Button();
            this.checkBoxRetrievePropOnlyIfVisible = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listBoxLog
            // 
            this.listBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBoxLog.FormattingEnabled = true;
            this.listBoxLog.Location = new System.Drawing.Point(9, 9);
            this.listBoxLog.Name = "listBoxLog";
            this.listBoxLog.Size = new System.Drawing.Size(290, 251);
            this.listBoxLog.TabIndex = 0;
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.Location = new System.Drawing.Point(440, 266);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(75, 23);
            this.buttonClose.TabIndex = 1;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // checkBoxAlwaysOnTop
            // 
            this.checkBoxAlwaysOnTop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBoxAlwaysOnTop.AutoSize = true;
            this.checkBoxAlwaysOnTop.Checked = true;
            this.checkBoxAlwaysOnTop.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxAlwaysOnTop.Location = new System.Drawing.Point(12, 270);
            this.checkBoxAlwaysOnTop.Name = "checkBoxAlwaysOnTop";
            this.checkBoxAlwaysOnTop.Size = new System.Drawing.Size(92, 17);
            this.checkBoxAlwaysOnTop.TabIndex = 2;
            this.checkBoxAlwaysOnTop.Text = "Always on top";
            this.checkBoxAlwaysOnTop.UseVisualStyleBackColor = true;
            this.checkBoxAlwaysOnTop.CheckedChanged += new System.EventHandler(this.checkBoxAlwaysOnTop_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.checkBoxRetrievePropOnlyIfVisible);
            this.groupBox1.Controls.Add(this.textBoxPropId);
            this.groupBox1.Controls.Add(this.checkBoxRetrievePropertyonSelectionChange);
            this.groupBox1.Location = new System.Drawing.Point(303, 9);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(216, 252);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Debug Options";
            // 
            // textBoxPropId
            // 
            this.textBoxPropId.Location = new System.Drawing.Point(22, 41);
            this.textBoxPropId.Name = "textBoxPropId";
            this.textBoxPropId.Size = new System.Drawing.Size(74, 20);
            this.textBoxPropId.TabIndex = 1;
            this.textBoxPropId.Text = "0x0037001F";
            this.textBoxPropId.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // checkBoxRetrievePropertyonSelectionChange
            // 
            this.checkBoxRetrievePropertyonSelectionChange.AutoSize = true;
            this.checkBoxRetrievePropertyonSelectionChange.Location = new System.Drawing.Point(5, 18);
            this.checkBoxRetrievePropertyonSelectionChange.Name = "checkBoxRetrievePropertyonSelectionChange";
            this.checkBoxRetrievePropertyonSelectionChange.Size = new System.Drawing.Size(209, 17);
            this.checkBoxRetrievePropertyonSelectionChange.TabIndex = 0;
            this.checkBoxRetrievePropertyonSelectionChange.Text = "Retrieve property on SelectionChange:";
            this.checkBoxRetrievePropertyonSelectionChange.UseVisualStyleBackColor = true;
            // 
            // buttonClear
            // 
            this.buttonClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClear.Location = new System.Drawing.Point(249, 266);
            this.buttonClear.Margin = new System.Windows.Forms.Padding(2);
            this.buttonClear.Name = "buttonClear";
            this.buttonClear.Size = new System.Drawing.Size(50, 23);
            this.buttonClear.TabIndex = 4;
            this.buttonClear.Text = "Clear";
            this.buttonClear.UseVisualStyleBackColor = true;
            this.buttonClear.Click += new System.EventHandler(this.buttonClear_Click);
            // 
            // checkBoxRetrievePropOnlyIfVisible
            // 
            this.checkBoxRetrievePropOnlyIfVisible.AutoSize = true;
            this.checkBoxRetrievePropOnlyIfVisible.Location = new System.Drawing.Point(102, 43);
            this.checkBoxRetrievePropOnlyIfVisible.Name = "checkBoxRetrievePropOnlyIfVisible";
            this.checkBoxRetrievePropOnlyIfVisible.Size = new System.Drawing.Size(109, 17);
            this.checkBoxRetrievePropOnlyIfVisible.TabIndex = 2;
            this.checkBoxRetrievePropOnlyIfVisible.Text = "Only if item visible";
            this.checkBoxRetrievePropOnlyIfVisible.UseVisualStyleBackColor = true;
            // 
            // FormEventTracker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(527, 294);
            this.Controls.Add(this.buttonClear);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.checkBoxAlwaysOnTop);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.listBoxLog);
            this.Name = "FormEventTracker";
            this.Text = "Outlook Events";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxLog;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.CheckBox checkBoxAlwaysOnTop;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.CheckBox checkBoxRetrievePropertyonSelectionChange;
        public System.Windows.Forms.TextBox textBoxPropId;
        private System.Windows.Forms.Button buttonClear;
        public System.Windows.Forms.CheckBox checkBoxRetrievePropOnlyIfVisible;
    }
}