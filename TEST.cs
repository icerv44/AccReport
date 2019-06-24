using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading.Tasks;





namespace Account_CheckReport
{
    public partial class TEST : Form
    {
        public TEST()
        {
            InitializeComponent();
        }

        private void TEST_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)

        {

            //SqlConnection cnn;
            SqlConnection myConnection = new SqlConnection("user id=user01;" +
                                       "password=user01;server=localhost;" +
                                       "Trusted_Connection=yes;" +
                                       "database=TEST; " +
                                       "connection timeout=30");
           /* string connetionString;
            
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            MessageBox.Show("Connection Open  !");
            cnn.Close();*/
            try
            {
                myConnection.Open();
                MessageBox.Show("Connection Open  !");
            }
            catch (Exception  er)
            {
                Console.WriteLine(er.ToString());
                MessageBox.Show(er.ToString());
            }


        }
    }
}
