using Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace report
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Excel | *.xls;*.xlsx;",
                Title = "Select File"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                dataGridView1.DataSource = importData(ofd.FileName);
            }
        }

        private DataView importData(string fileName)
        {
            String connection = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0;'", fileName);
            OleDbConnection connector = new OleDbConnection(connection);
            connector.Open();
            OleDbCommand query = new OleDbCommand("Select * from [DIVIPOLA-_C_digos_municipios$]", connector);
            OleDbDataAdapter adapter = new OleDbDataAdapter
            {
                SelectCommand = query
            };
            DataSet ds = new DataSet();
            adapter.Fill(ds);
            connector.Close();
            return ds.Tables[0].DefaultView;
        }

        private void cbFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
    }
}
