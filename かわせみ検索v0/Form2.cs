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
        class nx
        {
            //配列のサイズを固定で設定します。

            //最大テーブル大きさ
            public static int maxx = 230;　//項目数
            public static int maxy = 30;　//件数
            //訂正後テーブル大きさ
            public static int maxfx = 150;　//項目数
            public static int maxfy = 30;　//件数        
        }
    //読み込みボタンよりExcelから読み込みます
        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            //読み込みファイル名を指定します
            string ExcelBookFileName = @"C:\Users\淳一\Documents\Visual Studio 2013\Projects\かわせみ検索v0\かわせみ検索v0\bin\Debug\test.xlsx";

            //2次元配列ar[row,columun]を設定します。
            string[,] ar = new string[nx.maxx+20, nx.maxy+20];
            string[,] col = new string[nx.maxfx+10, 2];
            string[] skip_n = new string[]{
                "1",  //A 0
                "0",  //B 1
                "0",  //C 2
                "1",  //D 3
                "1",  //E 4
                "1",  //F 5
                "1",  //G 6 
                "1",  //H 7
                "1",  //I 8
                "0",  //J 9
                "0",  //K 10
                "0",  //L 11
                "0",  //M 12
                "0",  //N 13
                "0",  //O 14
                "0",  //P 15
                "0",  //Q 16
                "0",  //R 17
                "0",  //S 18
                "0",  //T 19
                "0",  //U 20
                "0",  //V 21
                "0",  //W 22
                "0",  //X 23
                "0",  //Y 24
                "0",  //Z 25
                "0",  //AA26
                "0",  //B 27
                "0",  //C 28
                "0",  //D 29
                "0",  //E 30
                "0",  //F 31
                "0",  //G 32
                "0",  //H 33
                "0",  //I 34
                "0",  //J 35
                "0",  //K 6
                "0",  //L 7
                "0",  //M 8
                "0",  //N 9
                "0",  //O 40
                "0",  //P 1
                "0",  //Q 2
                "0",  //R 3
                "0",  //S 4
                "0",  //T 5
                "0",  //U 6
                "0",  //V 7
                "0",  //W 8
                "0",  //X 9
                "0",  //Y 50
                "0",  //Z 1
                "0",  //BA2
                "0",  //B 3
                "0",  //C 4
                "0",  //D 5
                "0",  //E 6
                "0",  //F 7
                "0",  //G 8
                "0",  //H 9
                "0",  //I 60
                "0",  //J 1
                "0",  //K 2
                "0",  //L 3
                "0",  //M 4
                "0",  //N 5
                "0",  //O 6
                "0",  //P 7
                "0",  //Q 8
                "0",  //R 9
                "0",  //S 70
                "0",  //T 1
                "0",  //U 2
                "0",  //V 3
                "0",  //W 4
                "0",  //X 5
                "0",  //Y 6
                "0",  //Z 7
                "0",  //CA8
                "0",  //B 9
                "0",  //C 80
                "0",  //D 1
                "0",  //E 2
                "0",  //F 3
                "0",  //G 4
                "0",  //H 5
                "0",  //I 6
                "0",  //J 7
                "0",  //K 8
                "0",  //L 9
                "0",  //M 90
                "0",  //N 1
                "0",  //O 2
                "0",  //P 3
                "0",  //Q 4
                "0",  //R 5
                "0",  //S 6 
                "0",  //T 7
                "0",  //U 8
                "0",  //V 9
                "0",  //W 100
                "0",  //X 1
                "0",  //Y 2
                "0",  //Z 3
                "0",  //DA4
                "0",  //A 5
                "0",  //B 6
                "0",  //C 7
                "0",  //D 8
                "0",  //E 9
                "0",  //F 110
                "0",  //G 1
                "0",  //H 2
                "0",  //I 3
                "0",  //J 4
                "0",  //K 5
                "0",  //L 6
                "0",  //M 7
                "0",  //N 8
                "0",  //O 9
                "0",  //P 120
                "0",  //Q 1
                "0",  //R 2
                "0",  //S 3
                "0",  //T 4
                "0",  //U 5
                "0",  //V 6
                "0",  //W 7
                "0",  //X 8
                "0",  //Y 9
                "0",  //Z 130
                "0",  //EA1
                "0",  //A 2
                "0",  //B 3
                "0",  //C 4
                "0",  //D 5
                "0",  //E 6
                "0",  //F 7
                "0",  //G 8
                "0",  //H 9
                "0",  //I 140
                "0",  //J 1
                "0",  //K 2
                "0",  //L 3
                "0",  //M 4
                "0",  //N 5
                "0",  //O 6
                "0",  //P 7
                "1",  //Q 8
                "0",  //R 9
                "1",  //S 150
                "1",  //T 1
                "1",  //U 2
                "1",  //V 3
                "1",  //W 4
                "1",  //X 5
                "1",  //Y 6
                "1",  //Z 7
                "1",  //FA8
                "1",  //A 9
                "0",  //B 160
                "1",  //C 1
                "1",  //D 2
                "1",  //E 3
                "1",  //F 4
                "1",  //G 5
                "1",  //H 6
                "1",  //I 7
                "1",  //J 8
                "1",  //K 9
                "1",  //L 170
                "1",  //M 1
                "1",  //N 2
                "1",  //O 3
                "1",  //P 4
                "1",  //Q 5
                "1",  //R 6
                "1",  //S 7
                "0",  //T 8
                "1",  //U 9
                "1",  //V 180
                "1",  //W 1
                "1",  //X 2
                "1",  //Y 3
                "0",  //Z 4
                "0",  //GA5
                "1",  //B 6 
                "0",  //C 7
                "1",  //D 8
                "1",  //E 9
                "1",  //F 190
                "0",  //G 1
                "1",  //H 2
                "0",  //I 3
                "1",  //J 4
                "0",  //K 5
                "1",  //L 6
                "1",  //M 7
                "1",  //N 8
                "0",  //O 9
                "1",  //P 200
                "0",  //Q 1
                "0",  //R 2
                "1",  //S 3
                "0",  //T 4
                "1",  //U 5
                "1",  //V 6
                "0",  //W 7
                "0",  //X 8
                "0",  //Y 9
                "1",  //Z 210
                "0",  //HA1
                "1",  //B 2
                "0",  //C 3
                "1",  //D 4
                "1",  //E 5
                "1",  //F 6
                "1",  //G 7
                "1",  //H 8
                "1",  //I 9
                "1",  //J 220
                "1",  //K 1
                "0",  //L 2
                "0",  //M 3
                "0",  //N 4
                "0",  //O 5
                "0",  //P 6
                "0",  //Q 7
                "0",  //R 8
                "0",  //S 9
                "0",  //T 230
                "0",  //U 1
                "0",  //V 2
                "0",  //W 3
                "0",  //X 4
                "0",  //Y 5
                "0",  //Z 6
                };


            //Excelアプリケーションオブジェクトを作成します。アプリケーションウィンドウは非表示にします。
            ExcelApp.Visible = false;
            
            //Excelを読み込みます
            Workbook wb = ExcelApp.Workbooks.Open(ExcelBookFileName);

            //ExcelのSheet1を指定
            Worksheet ws1 = wb.Sheets[1];
            ws1.Select(Type.Missing);

            int fake_x = 0;
            int fake_y = 0;
            for (int x = 1; x < nx.maxx+2; x++)
                {
                        if (skip_n[x] == "0")
                        //非表示スイッチが0の場合
                        {
                        }
                        else
                        //非表示スイッチが1の場合
                        {
                            Range index = ws1.Cells[1, x];
                            dynamic val = index.Value2;
                            col[fake_x, 0] += "["+fake_x+","+x+"]"+Convert.ToString(val);
                            fake_x++; //スキップ列番号を増やす
                        }

                }
            

            for (int y = 2; y < nx.maxy; y++)
            {
                Range rgn = ws1.Cells[y, 1];
                dynamic val1 = rgn.Value2;
                String t_tex = Convert.ToString(val1);
                //ar[0, y] = y.ToString();
                //ar[1, y] = "*"+fake_y.ToString();

                if (t_tex == null)
                {
                //    なにもしない
                }
                else
                {
                fake_y++; //スキップ行番号を増やす
                fake_x = 0;
                   for (int x = 0; x < nx.maxx; x++)
                    {
                       if (skip_n[x] == "0")
                           //非表示スイッチが0の場合
                       {
                       }
                       else
                           //非表示スイッチが1の場合
                       {
                           Range index = ws1.Cells[y, x+1];
                           dynamic val = index.Value2; 
                           ar[fake_x, fake_y - 1] += Convert.ToString(val);
                           fake_x++; //スキップ列番号を増やす
                       }
                    }
                }

                
            }
            wb.Close(false); //ブッククローズ
            ExcelApp.Quit(); //Excel終了


      //データグリッドビューに代入



            int rim_y = fake_y + 0;//読み込み位置限度の設定

            label1.Text = nx.maxfy.ToString()+"件読み込み、有効"+rim_y+"件";





            dataGridView1.ColumnCount = fake_x+3; //gridviewの行・列設定
            dataGridView1.RowCount = fake_y;



            //DataGridView1の行ヘッダーに項目名を設定
            for (int i = 0; i < fake_x+1; i++)
            {
                dataGridView1.Columns[i].HeaderCell.Value = i.ToString()+ col[i, 0];
            }

            //行ヘッダーの幅を自動調節する
            dataGridView1.AutoResizeRowHeadersWidth(
                DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);


            for (int y = 0; y < rim_y; y++)
            {
            　　    for (int x = 0; x < fake_x+1; x++)
                    {
                    dataGridView1[x, y].Value =ar[x, y];//配列arをグリッドに並べる
                    }                     
                }
            }
           

        }

    }

