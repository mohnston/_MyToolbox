namespace Westmark
{
    partial class frmFindJob
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
            this.SearchStringTextBox = new System.Windows.Forms.TextBox();
            this.CloseFindJobButton = new System.Windows.Forms.Button();
            this.OKFindJobButton = new System.Windows.Forms.Button();
            this.SearchResultsListBox = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // SearchStringTextBox
            // 
            this.SearchStringTextBox.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SearchStringTextBox.Location = new System.Drawing.Point(0, 0);
            this.SearchStringTextBox.Name = "SearchStringTextBox";
            this.SearchStringTextBox.Size = new System.Drawing.Size(217, 31);
            this.SearchStringTextBox.TabIndex = 0;
            this.SearchStringTextBox.TextChanged += new System.EventHandler(this.SearchStringTextBox_TextChanged);
            this.SearchStringTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.SearchStringTextBox_KeyDown);
            this.SearchStringTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.SearchStringTextBox_KeyPress);
            // 
            // CloseFindJobButton
            // 
            this.CloseFindJobButton.BackColor = System.Drawing.Color.MistyRose;
            this.CloseFindJobButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CloseFindJobButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.CloseFindJobButton.Location = new System.Drawing.Point(260, 0);
            this.CloseFindJobButton.Name = "CloseFindJobButton";
            this.CloseFindJobButton.Size = new System.Drawing.Size(31, 31);
            this.CloseFindJobButton.TabIndex = 1;
            this.CloseFindJobButton.Text = "X";
            this.CloseFindJobButton.UseVisualStyleBackColor = false;
            // 
            // OKFindJobButton
            // 
            this.OKFindJobButton.BackColor = System.Drawing.Color.Honeydew;
            this.OKFindJobButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OKFindJobButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.OKFindJobButton.Location = new System.Drawing.Point(223, 0);
            this.OKFindJobButton.Name = "OKFindJobButton";
            this.OKFindJobButton.Size = new System.Drawing.Size(31, 31);
            this.OKFindJobButton.TabIndex = 2;
            this.OKFindJobButton.UseVisualStyleBackColor = false;
            this.OKFindJobButton.Click += new System.EventHandler(this.OKFindJobButton_Click);
            // 
            // SearchResultsListBox
            // 
            this.SearchResultsListBox.BackColor = System.Drawing.SystemColors.Window;
            this.SearchResultsListBox.Font = new System.Drawing.Font("Consolas", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SearchResultsListBox.FormattingEnabled = true;
            this.SearchResultsListBox.ItemHeight = 22;
            this.SearchResultsListBox.Location = new System.Drawing.Point(0, 31);
            this.SearchResultsListBox.Name = "SearchResultsListBox";
            this.SearchResultsListBox.Size = new System.Drawing.Size(291, 246);
            this.SearchResultsListBox.TabIndex = 3;
            // 
            // frmFindJob
            // 
            this.AcceptButton = this.OKFindJobButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CloseFindJobButton;
            this.ClientSize = new System.Drawing.Size(292, 279);
            this.Controls.Add(this.SearchResultsListBox);
            this.Controls.Add(this.OKFindJobButton);
            this.Controls.Add(this.CloseFindJobButton);
            this.Controls.Add(this.SearchStringTextBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmFindJob";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "frmFindJob";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button CloseFindJobButton;
        private System.Windows.Forms.Button OKFindJobButton;
        private System.Windows.Forms.ListBox SearchResultsListBox;
        public System.Windows.Forms.TextBox SearchStringTextBox;
    }
}