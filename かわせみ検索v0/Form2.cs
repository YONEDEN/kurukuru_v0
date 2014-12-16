using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace かわせみ検索v0
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();

        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            string ExcelBookFileName = @"C:\Users\淳一\Documents\Visual Studio 2013\Projects\かわせみ検索v0\かわせみ検索v0\bin\Debug\test.xlsx";
            string mestx = "";
            string[,] ar = new string[10, 10];
       //     Microsoft.Office.Interop.Excel.Application ExcelApp
       //       = new Microsoft.Office.Interop.Excel.Application();
            //Excelアプリケーションオブジェクトを作成します。アプリケーションウィンドウは非表示にします。
            ExcelApp.Visible = false;
            Workbook wb = ExcelApp.Workbooks.Open(ExcelBookFileName);
            //ExcelのSheet1を指定
            Worksheet ws1 = wb.Sheets[1];
            ws1.Select(Type.Missing);

            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    Range rgn = ws1.Cells[i, j];
                    dynamic val = rgn.Value2;
                    ar[i,j] += Convert.ToString(val) + "\r\n";
                }
            }
            wb.Close(false); //ブッククローズ
            ExcelApp.Quit(); //Excel終了


            String st_Result = "";
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    st_Result += ar[i,j];
                    st_Result += " ";
                }
                //st_Result += "\r\n";
            }
            for (int i = 1; i < 10; i++)
            {
                    dataGridView1.Columns.Add("ColItemCode", ar[1,i]);
            }
            
            
            
         //   MessageBox.Show(st_Result + "を取得しました");

       
       }

    }
}
