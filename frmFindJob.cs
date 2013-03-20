using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using sysData = System.Data;
using pgDBTools;

namespace Westmark
{
    public partial class frmFindJob : Form
    {
        public frmFindJob()
        {
            InitializeComponent();
        }
        public string selectedJobNo = string.Empty;

        private void SearchStringTextBox_TextChanged(object sender, EventArgs e)
        {
            if (SearchStringTextBox.Text.Length < 1) return;
            pgDB odb = new pgDB();
            string sql = "SELECT jobid, (jobid || '-' || name) as jobname " +
                "FROM job " + 
                "WHERE name ILIKE '%" + SearchStringTextBox.Text + "%' " +
                "ORDER BY jobid DESC";
            sysData.DataTable dt = odb.pgDataTable("eng", sql);
            SearchResultsListBox.DataSource = dt;
            SearchResultsListBox.DisplayMember = "jobname";
            SearchResultsListBox.ValueMember = "jobid";
        }

        private void OKFindJobButton_Click(object sender, EventArgs e)
        {
            if (SearchResultsListBox.SelectedValue != null)
            {
                selectedJobNo = SearchResultsListBox.SelectedValue.ToString();
            }
        }

        private void SearchStringTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void SearchStringTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                SearchResultsListBox.Focus();
                e.Handled = true;
            }
        }
    }
}
