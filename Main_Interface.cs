using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;                                                                                                                       
using System.Data.OleDb;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Globalization;
using MySql.Data.MySqlClient;
using System.Threading;
using Globalization = System.Globalization;
//using MySql.Data.MySqlClient;
/*
    Created by Chontee 
 
 */

namespace Account_CheckReport
{

    public partial class Main_Interface : Form
    {
        OleDbConnection bookConn;
        OleDbCommand oleDbCmd;
        OleDbDataReader mdr;
        //parameter from mdsaputra.udl
        String connParam = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\xt6000635\Documents\Invoice.mdb;Persist Security Info=True;User ID=admin";

       // static MySqlConnection con = new MySqlConnection("server = localhost; database = invoice; username = root; password=; ");
        /**********   LIST  ************/
        List<Processing.List_txt> lt_Exc = new List<Processing.List_txt>();
        List<Processing.List_txt> lt_Pdf = new List<Processing.List_txt>();
        List<Processing.List_txt> lt_Miss = new List<Processing.List_txt>();
       

        public Main_Interface()
        {
          //  bookConn = new OleDbConnection(connParam);
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //textBox2.Text = "Test";
            //Thread t2 = new Thread(new ThreadStart(SelectAll));
            //t2.IsBackground = true
            //t2.Start();
            SelectAll();
            textBox2.ScrollBars = ScrollBars.Both;

        }

        public  void SelectAll() {

            Select_Invoice_DNTH();
            Select_Invoice_SDM();
            Select_Invoice_SKD();
            Select_Invoice_ASTH();
            Select_Invoice_ANDEN();
        }

        public void Select_Invoice_DNTH()
        {
            string query = "SELECT * FROM invoice_dnth";

            bookConn = new OleDbConnection(connParam);
            bookConn.Open();
            try
            {
                OleDbDataAdapter dAdapter = new OleDbDataAdapter(query, connParam);
                DataTable dt = new DataTable();

                dAdapter.Fill(dt);
                dataGridView_DNTH.DataSource = dt;
            }
            catch (Exception er)
            {

                MessageBox.Show("ERROR : " + er);
                bookConn.Close();
            }
            bookConn.Close();
            
        }
        public  void Select_Invoice_SDM()
        {
            string query = "SELECT * FROM invoice_sdm";

            bookConn = new OleDbConnection(connParam);
            bookConn.Open();
            try
            {
                OleDbDataAdapter dAdapter = new OleDbDataAdapter(query, connParam);
                DataTable dt = new DataTable();

                dAdapter.Fill(dt);
                dataGridView_SDM.DataSource = dt;
            }
            catch (Exception er)
            {

                MessageBox.Show("ERROR : " + er);
                bookConn.Close();
            }
            bookConn.Close();
           
        }
        public  void Select_Invoice_SKD()
        {
            string query = "SELECT * FROM invoice_skd";

            bookConn = new OleDbConnection(connParam);
            bookConn.Open();
            try
            {
                OleDbDataAdapter dAdapter = new OleDbDataAdapter(query, connParam);
                DataTable dt = new DataTable();

                dAdapter.Fill(dt);
                dataGridView_SKD.DataSource = dt;
            }
            catch (Exception er)
            {

                MessageBox.Show("ERROR : " + er);
                bookConn.Close();
            }
            bookConn.Close();

            
        }
        public  void Select_Invoice_ASTH()
        {
            string query = "SELECT * FROM invoice_asth";

            bookConn = new OleDbConnection(connParam);
            bookConn.Open();
            try
            {
                OleDbDataAdapter dAdapter = new OleDbDataAdapter(query, connParam);
                DataTable dt = new DataTable();

                dAdapter.Fill(dt);
                dataGridView_ASTH.DataSource = dt;
            }
            catch (Exception er)
            {

                MessageBox.Show("ERROR : " + er);
                bookConn.Close();
            }
            bookConn.Close();

           
        }
        public  void Select_Invoice_ANDEN()
        {
            string query = "SELECT * FROM invoice_adth";

            bookConn = new OleDbConnection(connParam);
            bookConn.Open();
            try{
                OleDbDataAdapter dAdapter = new OleDbDataAdapter(query, connParam);
                DataTable dt = new DataTable();

                dAdapter.Fill(dt);
                dataGridView_ANDEN.DataSource = dt;
            }catch(Exception er){

                MessageBox.Show("ERROR : " + er);
                bookConn.Close();
            }
            bookConn.Close();
        }

        public  void Insert_Invoice(int Year, int InvNo)
        {
            string query = "INSERT INTO invoice_" + comboBox1.Text + " ( [Year], [Inv_No]) VALUES (" + Year + "," + InvNo + ")";

            bookConn = new OleDbConnection(connParam);
            oleDbCmd = new OleDbCommand(query, bookConn);
            bookConn.Open();
            try {


                //oleDbCmd.Parameters.AddWithValue("@Year", Year.ToString());
                //oleDbCmd.Parameters.AddWithValue("@Inv_No", InvNo);
                oleDbCmd.ExecuteNonQuery();

               
            }catch(Exception er){
                MessageBox.Show("ERROR : " + er);
                bookConn.Close();
            }
            bookConn.Close();
        }


        private  void button3_Click(object sender, EventArgs e)
        {
            //Thread workThread = new Thread(new ThreadStart(excelInsert));
            //workThread.IsBackground = true;
            //workThread.Start();
            //excelInsert();

            excelInsert();

        }
        public void DISTINCTyear(string company) {

            string query = "SELECT DISTINCT Year FROM Invoice_" + company+"";

            
            bookConn = new OleDbConnection(connParam);
            bookConn.Open();
            try
            {
                oleDbCmd = new OleDbCommand(query, bookConn);

                mdr = oleDbCmd.ExecuteReader();

                while (mdr.Read())
                {
                    Date_1.Items.Add(mdr.GetInt32(0));
                    Date_2.Items.Add(mdr.GetInt32(0));
                }
            }
            catch (Exception er) {

                MessageBox.Show("ERROR : " + er);
               bookConn.Close();
            }
           bookConn.Close();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Date_1.Items.Clear();
            Date_2.Items.Clear();
           string  select =  comboBox1.Text;
           DISTINCTyear(select);
        }

        string Path_DNTH_Invoice = @"\\xts02\DIAT\00.COMMON\00.Special Folder\COMN_JETFORM\ACC\INVOICE";
        string Path_DNTH_DNCN = @"\\xts02\DIAT\00.COMMON\00.Special Folder\COMN_JETFORM\ACC\DNCN";
        string Path_DNTH_Spart = @"\\xts02\DIAT\00.COMMON\00.Special Folder\COMN_JETFORM\ACC\INVSPART";
        string Path_SDM_Invoice = @"\\xts06\SDM\00.COMMON_SDM\00.Special_Folder\COMN_JETFORM\INVOICE";
        string Path_SDM_DNCN = @"\\xts06\SDM\00.COMMON_SDM\00.Special_Folder\COMN_JETFORM\DNCN";
        string Path_SDM_Spart = @"\\xts06\SDM\00.COMMON_SDM\00.Special_Folder\COMN_JETFORM\PDF\INVOICE\INVSPART";
        string Path_SKD_Invoice = @"\\xts06\SKD_New\00.COMMON_SKD\00.Special_Folder\COMN_JETFORM\PDF\INVOICE";
        string Path_SKD_DNCN = @"\\xts06\SKD_New\00.COMMON_SKD\00.Special_Folder\COMN_JETFORM\DNCN";
        string Path_SKD_Spart = @"\\xts06\SKD_New\00.COMMON_SKD\00.Special_Folder\COMN_JETFORM\PDF\INVOICE\INVSPART";
        string Path_ASTH_Invoice = @"\\xts02\DIAT\00.COMMON\00.Special Folder\COMN_JETFORM\ACC\ASTH\INVOICE";
        string Path_ASTH_DNCN = @"\\xts02\DIAT\00.COMMON\00.Special Folder\COMN_JETFORM\ACC\ASTH\DNCN";
        string Path_ASTH_Spart = @"\\xts02\DIAT\00.COMMON\00.Special Folder\COMN_JETFORM\ACC\ASTH\INVSPART";
        string Path_ADTH_Invoice = @"\\xts15\DNTH\00.COMMON\00.Special Folder\COMN_JETFORM\User_Data\COMN_JETFORM\INVOICE";
        string Path_ADTH_DNCN = @"\\xts15\DNTH\00.COMMON\00.Special Folder\COMN_JETFORM\DNCN";
        string Path_ADTH_Spart = @"\\xts15\DNTH\00.COMMON\00.Special Folder\COMN_JETFORM\User_Data\COMN_JETFORM\INVSPART";




        public void YearSearch(int start ,int end )
        {
           
            DateTime startDate = new DateTime();
           
            DateTime endDate = new DateTime();
            
            Globalization.CultureInfo customCulture = new Globalization.CultureInfo("en-US", true);

            DateTime.TryParseExact(start.ToString(), "yyyyMM",
                   customCulture,
                   DateTimeStyles.None, out startDate);
                          
              DateTime.TryParseExact(end.ToString(), "yyyyMM",
                   customCulture,
                   DateTimeStyles.None, out endDate);
              

              for (DateTime i = startDate; i <= endDate; i = i.AddMonths(1)) {

                  string tmp = i.ToString("yyyyMM", customCulture);
                  Select_Year_Search(Int32.Parse(tmp));
                  //Console.WriteLine("Date : " + i);
              }

        }

        public void Select_Year_Search(int year) {


            string query = "SELECT DISTINCT Inv_No, Year FROM Invoice_" + comboBox1.Text + "  WHERE  Year = " + year + " and Inv_No <= 900000 and (Inv_No <=600000 or Inv_No >=700000)";


            bookConn = new OleDbConnection(connParam);
            bookConn.Open();
            label3.Text = year.ToString();
            try
            {
                oleDbCmd = new OleDbCommand(query, bookConn);
                int rowCount = (int)oleDbCmd.ExecuteScalar();
                
 
                mdr = oleDbCmd.ExecuteReader();
               


                while (mdr.Read())
                {

                    Diretory(mdr.GetInt32(0), mdr.GetInt32(1));
                   
                   // Date_1.Items.Add(mdr.GetInt32(0));
                    
                }
            }
            catch (Exception er)
            {

                MessageBox.Show("ERROR : " + er);
                bookConn.Close();
            }
            bookConn.Close();
        
        }

        public void Diretory( int  Inv,int year) {
             string invPath = "";
            string dncnPath = "";
            string sPartPath = "";
            
            switch (comboBox1.Text)
            {
                case "DNTH":
                    invPath = Path_DNTH_Invoice; 
                    dncnPath =  Path_DNTH_DNCN ;
                    sPartPath = Path_DNTH_Spart;
                    break;
                case "SDM":
                     invPath = Path_SDM_Invoice;
                     dncnPath = Path_SDM_DNCN;
                     sPartPath = Path_SDM_Spart;   
                    break;
                case "SKD":
                    invPath = Path_SKD_Invoice;
                    dncnPath = Path_SKD_DNCN;
                    sPartPath = Path_SKD_Spart;
                       break;
                case "ASTH":
                    invPath = Path_ASTH_Invoice;
                    dncnPath = Path_ASTH_DNCN;
                    sPartPath = Path_ASTH_Spart;
                    break;
                case "ADTH":
                    invPath = Path_ADTH_Invoice;
                    dncnPath = Path_ADTH_DNCN;
                    sPartPath = Path_ADTH_Spart;
                    break;
                default:
                    Console.WriteLine("Invalid Company");
                    MessageBox.Show("Invalid Company");
                    break;
            }
            
             
            string[] filesInv = Directory.GetFiles(invPath + "\\" + year, Inv.ToString("000000") + ".pdf", SearchOption.AllDirectories);
            string[] filesDncn = Directory.GetFiles(dncnPath + "\\" + year, Inv.ToString("000000") + ".pdf", SearchOption.AllDirectories);
            string[] filesSpart = Directory.GetFiles(sPartPath + "\\" + year, Inv.ToString("000000") + ".pdf", SearchOption.AllDirectories);
            
            if (filesInv.Length == 0 & filesDncn.Length == 0 & filesSpart.Length == 0)
                {
                    Processing.List_txt ltMiss = new Processing.List_txt();
                    ltMiss.MissYear = year.ToString();
                    ltMiss.MissReportNo = Inv.ToString("000000");
                    lt_Miss.Add(ltMiss);


                      
                    //textBox2.TextChanged();
                   textBox2.AppendText("Missing Year = " + year.ToString() + "  Inv No = " + Inv.ToString("000000") + "\n");
                   
                }
            
            
            
           
        }

        public  void excelInsert()
        {

            //OpenFileDialog openFileDialog1 = new OpenFileDialog
            //{
            //    InitialDirectory = @"D:\",
            //    Title = "Browse Excel Files",

            //    CheckFileExists = true,
            //    CheckPathExists = true,
            //    DefaultExt = "xls",
            //    Filter = "Excel files (*.xls/*.xlsx)|*.xls;*xlsx",
            //    FilterIndex = 2,
            //    RestoreDirectory = true,

            //    ReadOnlyChecked = true,
            //    ShowReadOnly = true
            //};
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
               // textBox1.Text = openFileDialog1.FileName;
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
                SelectAll();
                MessageBox.Show("INSERT Finish");

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
            catch (Exception er) {
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

        
        
        }

      
    






        //int cout = 0;
        //bool flag_Master = false;
        /********************************/


        /*******************************************************************/
        public void ClearValue() {
            Date_1.Items.Clear();
            Date_2.Items.Clear();
            //Path1.Clear();
            //Path2.Clear();
            //Path3.Clear();
            //Path4.Clear();
            //Path5.Clear();
            //textBox2.Clear();
            lt_Miss.Clear();
            lt_Exc.Clear();
            lt_Pdf.Clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            YearSearch(Int32.Parse(Date_1.Text),Int32.Parse(Date_2.Text));
            Int32 length = lt_Miss.Count;
            TotalMissRe.Text = length.ToString();
            lt_Miss.Clear();
        }
       







        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
          
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private  void dataGridView_DNTH_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        
        private void vScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
           
        }

     
        

    
    }
}
