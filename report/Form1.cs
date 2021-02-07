using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace report
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public const string FILTER = "Nombre Departamento";
        public const string TYPE = "Tipo: Municipio / Isla / Área no municipalizada";
        private DataSet ds;

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
            DataTable dt = ds.Tables[0];
            string department = "";
            foreach (DataRow row in dt.Rows)
            {
                if (department == (string)row[FILTER])
                    department = (string)row[FILTER];
                else
                {
                    cbFilter.Items.Add(row[FILTER]);
                    department = (string)row[FILTER];
                }
            }
            generateChart();
        }

        private void generateChart()
        {
            DataTable dt = ds.Tables[0];
            int municipios = 0, islas = 0, noMun = 0;
            foreach (DataRow row in dt.Rows)
            {
                if ("Municipio" == (string)row[TYPE])
                    municipios++;
                else if ("Isla" == (string)row[TYPE])
                    islas++;
                else
                    noMun++;
            }
            chart1.Series["Series1"].Points.AddXY("Municipio", municipios);
            chart1.Series["Series1"].Points.AddXY("Isla", islas);
            chart1.Series["Series1"].Points.AddXY("Área no municipalizada", noMun);
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
            ds = new DataSet();
            adapter.Fill(ds);
            connector.Close();
            return ds.Tables[0].DefaultView;
        }

        private void cbFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = ds.Tables[0];
            string department = cbFilter.GetItemText(cbFilter.SelectedItem);
            dt.DefaultView.RowFilter = string.Format("[{0}] LIKE '%{1}%'", FILTER, department);
        }
    }
}
