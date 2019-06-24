using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Account_CheckReport
{
    class Bin_TMP
    {
        /*******************************************************************/

        //public void button2(object sender, EventArgs e)
        //{
        //    ClearValue();


        //    OpenFileDialog openFileDialog1 = new OpenFileDialog
        //    {
        //        InitialDirectory = @"D:\",
        //        Title = "Browse Excel Files",

        //        CheckFileExists = true,
        //        CheckPathExists = true,
        //        DefaultExt = "xls",
        //        Filter = "Excel files (*.xls/*.xlsx)|*.xls;*xlsx",
        //        FilterIndex = 2,
        //        RestoreDirectory = true,

        //        ReadOnlyChecked = true,
        //        ShowReadOnly = true
        //    };

        //    if (openFileDialog1.ShowDialog() == DialogResult.OK)
        //    {
        //        textBox1.Text = openFileDialog1.FileName;
        //    }


        //    //********************************************************************************


        //    //List<Processing.List_txt> lt_Exc = new List<Processing.List_txt>();



        //    //List<string> lt_Year = new List<string>();
        //   // List<string> lt_textbox_year = new List<string>();


        //    Excel.Application xlApp = new Excel.Application();
        //    xlApp.DisplayAlerts = false;



        //    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@openFileDialog1.FileName);
        //    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //    Console.WriteLine(xlWorkbook.Sheets[1].name);
        //    Excel.Range xlRange = xlWorksheet.UsedRange;

        //    try
        //    {

        //        /*************************************************************************************/
        //        float Pro  ;

        //        progressBar1.Minimum = 0;
        //        progressBar1.Maximum = 100;


        //        /**************************************************************************************/
        //        int rowCount = xlRange.Rows.Count;
        //        int colCount = xlRange.Columns.Count;




        //        //int Year = 0;
        //        string p = openFileDialog1.FileName;
        //        textBox2_TextChanged(p+"\r\n");
        //        string squ ;
        //        for (int z = 2; z <= rowCount; z++)
        //        {


        //            Processing.List_txt lt = new Processing.List_txt();



        //            lt.Year = xlRange.Cells[z, 1].Value2.ToString();
        //            lt.ReportNo = xlRange.Cells[z, 2].Value2.ToString();
        //            lt.ReportNo = lt.ReportNo.PadLeft(6, '0');
        //            lt_Exc.Add(lt);


        //            squ = (z - 1) + ")";
        //            squ = squ.PadRight(5,'_');

        //             string tmp = "" ; // index year-1
        //            if(z >=3 ){
        //                tmp = lt_Exc[z - 3].Year;

        //            }

        //            if (lt_Exc[z - 2].Year != tmp)
        //            {
        //                Date_1.Items.Add(tmp);
        //                Date_2.Items.Add(tmp);

        //            }
        //            else if (z == rowCount) 
        //            {
        //                Date_1.Items.Add(lt_Exc[z - 2].Year);
        //                Date_2.Items.Add(lt_Exc[z - 2].Year);

        //            }

        //            textBox2_TextChanged(squ + lt_Exc[z-2].ToString());
        //            Pro = (((float)z / (float)rowCount ) * 100);
        //            int tmp_int = (int)Pro;
        //            progressBar1.Value = tmp_int;
        //            Percent_T.Text = tmp_int.ToString() + "%";
        //            Percent_T.Update();            
        //        }

        //        //rule of thumb for releasing com objects:
        //        //  never use two dots, all COM objects must be referenced and released individually
        //        //  ex: [somthing].[something].[something] is bad

        //        //release com objects to fully kill excel process from running in the background

        //        Marshal.ReleaseComObject(xlRange);
        //        Marshal.ReleaseComObject(xlWorksheet);

        //        //close and release
        //        xlWorkbook.Close();
        //        Marshal.ReleaseComObject(xlWorkbook);


        //        //quit and release

        //        xlApp.Quit();

        //        Marshal.ReleaseComObject(xlApp);

        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //    }
        //    catch (Exception er)
        //    {
        //        MessageBox.Show("Error : " + er);


        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();

        //        //rule of thumb for releasing com objects:
        //        //  never use two dots, all COM objects must be referenced and released individually
        //        //  ex: [somthing].[something].[something] is bad

        //        //release com objects to fully kill excel process from running in the background

        //        Marshal.ReleaseComObject(xlRange);
        //        Marshal.ReleaseComObject(xlWorksheet);

        //        //close and release
        //        xlWorkbook.Close();
        //        Marshal.ReleaseComObject(xlWorkbook);


        //        //quit and release

        //        xlApp.Quit();

        //        Marshal.ReleaseComObject(xlApp);
        //    }



        //    flag_Master = true;

        //}

        //public void Submit_Box(object sender, EventArgs e)
        //{
        //    float Pro;

        //    progressBar1.Minimum = 0;
        //    progressBar1.Maximum = 100;

        //    textBox2.Clear();
        //    lt_Pdf.Clear();
        //    string StartYear = Date_1.Text.Substring(0,4);
        //    string StartMonth = Date_1.Text.Substring(4, 2);

        //    string EndYear = Date_2.Text.Substring(0,4);
        //    string EndMonth = Date_2.Text.Substring(4, 2);

        //    int St_Year = int.Parse(StartYear);
        //    int St_Month = int.Parse(StartMonth);

        //    int En_Year = int.Parse(EndYear);
        //    int En_Month = int.Parse(EndMonth);

        //    DateTime Start_DT = new DateTime(St_Year,St_Month,1);



        //    DateTime End_DT = new DateTime(En_Year,En_Month,1);

        //    String Path_Txt = @"D:\ChkMissReport\ChkMissReport" + DateTime.Now.Year + DateTime.Now.Month + ".txt";

        //    List<String> path_str = new List<string>();
        //    //string[] path_Str ;D:\ChkMissReport

        //    string Path1_tmp = ((@"" + Path1.Text) == "" ? "0" : (@"" + Path1.Text));
        //    string Path2_tmp = ((@"" + Path2.Text) == "" ? "0" : (@"" + Path2.Text));
        //    string Path3_tmp = ((@"" + Path3.Text) == "" ? "0" : (@"" + Path3.Text));
        //    string Path4_tmp = ((@"" + Path4.Text) == "" ? "0" : (@"" + Path4.Text));
        //    string Path5_tmp = ((@"" + Path5.Text) == "" ? "0" : (@"" + Path5.Text));


        //     path_str.Add(Path1_tmp);
        //     path_str.Add(Path2_tmp);
        //     path_str.Add(Path3_tmp);
        //     path_str.Add(Path4_tmp);
        //     path_str.Add(Path5_tmp);




        //    int cout_Excel = 0;
        //    while (Start_DT <= End_DT)
        //    {

        //        string yymm = Start_DT.ToString("yyyyMM", CultureInfo.InvariantCulture);

        //        if (flag_Master != false) {

        //            for (int i = 0; i < lt_Exc.Count; i++)
        //            {


        //                if (lt_Exc[i].Year == yymm)
        //                {

        //                    cout_Excel++;
        //                }

        //            }

        //        }




        //      textBox2_TextChanged(@"\INVOICE\" + yymm + "\r\n");




        //      for (int i = 0; i < path_str.Count; i++)
        //      {

        //          Path1_box(path_str[i], yymm);

        //      }







        //    Start_DT = Start_DT.AddMonths(1);
        //    }

        //    bool miss = false;

        //    File.WriteAllText(Path_Txt, "*************************Missing Report************************\r\n");
        //    textBox2.Clear();
        //    int round = 1;
        //    foreach (var lt in lt_Exc) {
        //        int YY = int.Parse(lt.Year.Substring(0, 4));
        //        int MM = int.Parse(lt.Year.Substring(4, 2));

        //        DateTime Year_Excel = new DateTime(YY, MM, 1);


        //        if (Year_Excel >= End_DT.AddMonths(1)){
        //            int tmpp_int = 100;
        //            progressBar1.Value = tmpp_int;
        //            Percent_T.Text = tmpp_int.ToString() + "%";
        //            break;
        //        }

        //       foreach(var ltpdf in lt_Pdf)
        //       {
        //           miss = false;
        //           if (lt.ReportNo.ToString() == ltpdf.PdfNo.ToString() && lt.Year == ltpdf.PdfYear)
        //           {
        //               miss = false;
        //               break;
        //           }
        //           else if (lt.ReportNo != ltpdf.PdfNo  )
        //           {
        //               miss = true;

        //           }




        //       }
        //       Pro = (((float)round / (float)lt_Exc.Count) * 100);
        //       int tmp_int = (int)Pro;
        //       progressBar1.Value = tmp_int;
        //       Percent_T.Text = tmp_int.ToString() + "%";
        //       Percent_T.Update();

        //       round++;
        //       if (miss == true && Year_Excel <= End_DT)
        //       {

        //           Processing.List_txt ltMiss = new Processing.List_txt();

        //           ltMiss.MissReportNo = lt.ReportNo;
        //           ltMiss.MissYear = lt.Year;
        //           File.AppendAllText(Path_Txt, "Year : " + lt.Year + "  Miss Report No : " + lt.ReportNo +"\r\n");

        //           lt_Miss.Add(ltMiss);

        //           textBox2_TextChanged("Year : " + lt.Year + "    Miss No : " + lt.ReportNo + "\r\n");

        //       }




        //       miss = false;


        //    }




        //     TotalFileExcel.Text = cout_Excel.ToString()  ;
        //     TotalFilePdf.Text = lt_Pdf.Count.ToString() ;
        //     TotalMissRe.Text = lt_Miss.Count.ToString();
        //     TotalFileExcel.Update();
        //     TotalFilePdf.Update();
        //     TotalMissRe.Update();
        //     lt_Miss.Clear();

        //  /*  foreach (var lt in lt_Exc) {

        //        textBox2_TextChanged(lt.Year +"***" +lt.ReportNo +"\r\n");
        //    }*/


        //}


        //public void Path1_box(string str, string yymm)
        //{
        //    if (str != "0")
        //    {

        //        DirectoryInfo d = new DirectoryInfo(str + "\\" + yymm + "\\");//Assuming Test is your Folder
        //        FileInfo[] Files = d.GetFiles("*.pdf"); //Getting Text files         
        //        foreach (FileInfo file in Files)
        //        {
        //            Processing.List_txt ltPdf = new Processing.List_txt();
        //            //str = file.Name.Substring(0, file.Name.IndexOf(".")) + "\r\n";
        //            //ltPdf.Add(file.Name.Substring(0, file.Name.IndexOf(".")));
        //            ltPdf.PdfYear = yymm;
        //            ltPdf.PdfNo = file.Name.Substring(0, file.Name.IndexOf("."));
        //            lt_Pdf.Add(ltPdf);

        //            //File.AppendAllText(path, str);

        //            //Console.WriteLine(String.Join("\n", lt[cout]));
        //            //Console.WriteLine(str);
        //            cout++;
        //        }




        //    }
        //    else { 





        //        }


        //}




        //public void textBox2_TextChanged(string str)
        //{


        //    textBox2.ScrollBars = ScrollBars.Both;
        //    textBox2.Text += str;

        //}

    }
}
