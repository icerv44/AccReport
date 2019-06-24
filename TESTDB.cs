using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace Account_CheckReport
{
    public partial class TESTDB : Form
    {
        private OleDbConnection bookConn;
        private OleDbCommand oleDbCmd = new OleDbCommand();
        //parameter from mdsaputra.udl
        private String connParam = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\xt6000635\Documents\Invoice.mdb;Persist Security Info=True;User ID=admin";
        public TESTDB()
        {
            bookConn = new OleDbConnection(connParam);
            InitializeComponent();
        }

        private void TESTDB_Load(object sender, EventArgs e)
        {

        }
        public void Insert_Invoice(int Year, int InvNo)
        {
            try
            {
                string query = "INSERT INTO `invoice_DNTH`( `Year`, `Inv_No`) VALUES (" + Year + "," + InvNo + ")";

                OleDbDataAdapter dAdapter = new OleDbDataAdapter(query, connParam);

                DataTable dt = new DataTable();
                DataSet ds = new DataSet();

                dAdapter.Fill(dt);

                //dataGridView1.DataSource = dt;
            }
            catch (Exception er) {

                MessageBox.Show("ERROR : " + er);
            
            }
           

        }
        public void SelectQRY() {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            try
            {
                OleDbDataAdapter dAdapter = new OleDbDataAdapter("select * from Invoice_DNTH", connParam);
                OleDbCommandBuilder cBuilder = new OleDbCommandBuilder(dAdapter);

                DataTable dataTable = new DataTable();
                DataSet ds = new DataSet();

                dAdapter.Fill(dataTable);

                dataGridView1.DataSource = dataTable;

            }
            catch (Exception er)
            {
                bookConn.Close();
                MessageBox.Show("ERROR : " + er);
            }

            bookConn.Close();

        
        }
        public void excelInsert()
        {

            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Excel Files",

                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xls",
                Filter = "Excel files (*.xls/*.xlsx)|*.xls;*xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //textBox1.Text = openFileDialog1.FileName;
            }

            Excel.Application xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;



            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@openFileDialog1.FileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Console.WriteLine(xlWorkbook.Sheets[1].name);
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            try
            {
                for (int z = 2; z <= rowCount; z++)
                {


                    Processing.List_txt lt = new Processing.List_txt();



                    lt.Year = xlRange.Cells[z, 1].Value2.ToString();
                    string tmpY = "";
                    string tmpNo = "";
                    tmpY = lt.Year.Substring(0, 6);
                    lt.ReportNo = xlRange.Cells[z, 2].Value2.ToString();
                    tmpNo = lt.ReportNo;
                    Insert_Invoice(Int32.Parse(tmpY), Int32.Parse(tmpNo));


                }
                //SelectAll();
                MessageBox.Show("INSERT Finish");

            }
            catch (Exception er)
            {
                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                MessageBox.Show("Error : " + er);
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);


                //quit and release

                xlApp.Quit();

                Marshal.ReleaseComObject(xlApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();


            }

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);


            //quit and release

            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();


        }

        public void PDFSearch() { 
        
        
        
        
        
        }




        private void button1_Click(object sender, EventArgs e)
        {

            excelInsert();
        }
    }
}
