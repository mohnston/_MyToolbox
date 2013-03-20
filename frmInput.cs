using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace ProjectSpec
{
	/// <summary>
	/// Summary description for frmInput.
	/// </summary>
	public class frmInput : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.TextBox tbxValue;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Label lblPrompt;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmInput()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.btnOK = new System.Windows.Forms.Button();
            this.tbxValue = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblPrompt = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.BackColor = System.Drawing.Color.Honeydew;
            this.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Location = new System.Drawing.Point(259, 9);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(144, 37);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // tbxValue
            // 
            this.tbxValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbxValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbxValue.Location = new System.Drawing.Point(10, 111);
            this.tbxValue.Name = "tbxValue";
            this.tbxValue.Size = new System.Drawing.Size(393, 26);
            this.tbxValue.TabIndex = 2;
            this.tbxValue.Text = "textBox1";
            this.tbxValue.TextChanged += new System.EventHandler(this.tbxValue_TextChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.MistyRose;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(259, 55);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(144, 37);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblPrompt
            // 
            this.lblPrompt.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblPrompt.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrompt.Location = new System.Drawing.Point(10, 9);
            this.lblPrompt.Name = "lblPrompt";
            this.lblPrompt.Size = new System.Drawing.Size(240, 83);
            this.lblPrompt.TabIndex = 3;
            this.lblPrompt.Text = "Prompt";
            // 
            // frmInput
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(409, 142);
            this.Controls.Add(this.lblPrompt);
            this.Controls.Add(this.tbxValue);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.Name = "frmInput";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Input Box";
            this.Load += new System.EventHandler(this.frmInput_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion
		private string myCaption = "";
		private string myPrompt = "";
		private string myDefault = "";
		private string myInput = "";

		private void frmInput_Load(object sender, System.EventArgs e)
		{
			this.Text = this.Caption;
			lblPrompt.Text = this.Prompt;
			tbxValue.Text = this.DefaultValue;
			
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
            this.DialogResult = DialogResult.OK;
            myInput = tbxValue.Text;
			this.Hide();
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			myInput = "";
			this.Hide();
		}

		private void tbxValue_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		#region Properties
		public string Input 
		{
			get 
			{
				return myInput;
			}
		}

		public string Caption 
		{
			get 
			{
				return myCaption;
			}
			set 
			{
				myCaption = value;
			}
		}

		public string Prompt 
		{
			get 
			{
				return myPrompt;
			}
			set 
			{
				myPrompt = value;
			}
		}
		
		public string DefaultValue 
		{
			get 
			{
				return myDefault;
			}
			set 
			{
				myDefault = value;
			}
		}

		#endregion
	}
}
