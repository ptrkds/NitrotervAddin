namespace NitrotervOutlookAddIn
{
    partial class iktatasDialogForm
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
            this.saveButton = new System.Windows.Forms.Button();
            this.resetButton = new System.Windows.Forms.Button();
            this.localPathLabel = new System.Windows.Forms.Label();
            this.networkPathLabel = new System.Windows.Forms.Label();
            this.backButton = new System.Windows.Forms.Button();
            this.networkPathTextBox = new System.Windows.Forms.TextBox();
            this.localPathTextBox = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.folderBrowserDialog2 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // saveButton
            // 
            this.saveButton.Location = new System.Drawing.Point(12, 116);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(75, 23);
            this.saveButton.TabIndex = 0;
            this.saveButton.Text = "Mentés";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // resetButton
            // 
            this.resetButton.Location = new System.Drawing.Point(124, 116);
            this.resetButton.Name = "resetButton";
            this.resetButton.Size = new System.Drawing.Size(75, 23);
            this.resetButton.TabIndex = 1;
            this.resetButton.Text = "Alaphelyzet";
            this.resetButton.UseVisualStyleBackColor = true;
            this.resetButton.Click += new System.EventHandler(this.resetButton_Click);
            // 
            // localPathLabel
            // 
            this.localPathLabel.AutoSize = true;
            this.localPathLabel.Location = new System.Drawing.Point(13, 28);
            this.localPathLabel.Name = "localPathLabel";
            this.localPathLabel.Size = new System.Drawing.Size(103, 13);
            this.localPathLabel.TabIndex = 2;
            this.localPathLabel.Text = "Lokális mappa helye";
            // 
            // networkPathLabel
            // 
            this.networkPathLabel.AutoSize = true;
            this.networkPathLabel.Location = new System.Drawing.Point(13, 72);
            this.networkPathLabel.Name = "networkPathLabel";
            this.networkPathLabel.Size = new System.Drawing.Size(108, 13);
            this.networkPathLabel.TabIndex = 3;
            this.networkPathLabel.Text = "Hálózati mappa helye";
            // 
            // backButton
            // 
            this.backButton.Location = new System.Drawing.Point(239, 116);
            this.backButton.Name = "backButton";
            this.backButton.Size = new System.Drawing.Size(75, 23);
            this.backButton.TabIndex = 6;
            this.backButton.Text = "Mégse";
            this.backButton.UseVisualStyleBackColor = true;
            this.backButton.Click += new System.EventHandler(this.backButton_Click);
            // 
            // networkPathTextBox
            // 
            this.networkPathTextBox.Location = new System.Drawing.Point(139, 69);
            this.networkPathTextBox.Name = "networkPathTextBox";
            this.networkPathTextBox.ReadOnly = true;
            this.networkPathTextBox.Size = new System.Drawing.Size(175, 20);
            this.networkPathTextBox.TabIndex = 5;
            this.networkPathTextBox.Click += new System.EventHandler(this.networkPathTextBox_Click);
            // 
            // localPathTextBox
            // 
            this.localPathTextBox.Location = new System.Drawing.Point(139, 25);
            this.localPathTextBox.Name = "localPathTextBox";
            this.localPathTextBox.ReadOnly = true;
            this.localPathTextBox.Size = new System.Drawing.Size(175, 20);
            this.localPathTextBox.TabIndex = 4;
            this.localPathTextBox.Click += new System.EventHandler(this.localPathTextBox_Click);
            // 
            // iktatasDialogForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(333, 154);
            this.Controls.Add(this.backButton);
            this.Controls.Add(this.networkPathTextBox);
            this.Controls.Add(this.localPathTextBox);
            this.Controls.Add(this.networkPathLabel);
            this.Controls.Add(this.localPathLabel);
            this.Controls.Add(this.resetButton);
            this.Controls.Add(this.saveButton);
            this.Name = "iktatasDialogForm";
            this.ShowIcon = false;
            this.Text = "Beállítások";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.Button resetButton;
        private System.Windows.Forms.Label localPathLabel;
        private System.Windows.Forms.Label networkPathLabel;
        private System.Windows.Forms.Button backButton;
        private System.Windows.Forms.TextBox networkPathTextBox;
        private System.Windows.Forms.TextBox localPathTextBox;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog2;
    }
}