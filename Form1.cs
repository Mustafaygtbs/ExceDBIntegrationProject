using System.Collections;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExceDBIntegrationProject
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SqlConnection baglanti = new SqlConnection(@"Data Source=MUSTAFA\SQLEXPRESS;Initial Catalog=ProjelerVT;Integrated Security=True");
        private void btnVTdenOku_Click(object sender, EventArgs e)
        {
            Excel.Application excelUygulama = new Excel.Application();
            excelUygulama.Visible = true;
            Excel.Workbook wb = excelUygulama.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sayfa1 = (Excel.Worksheet)wb.Sheets[1];
            string[] basliklar = { "Personel no", "Ad", "Soyad", "Semt", "Sehir" };
            Excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = sayfa1.Cells[1, (1 + i)];
                range.Value2 = basliklar[i];
            }
            try
            {
                baglanti.Open();
                string sqlCumlesi = "SELECT PersonelNo, Ad, Soyad, Semt, Sehir FROM Personel";
                SqlCommand sqlCommand = new SqlCommand(sqlCumlesi, baglanti);
                SqlDataReader sdr = sqlCommand.ExecuteReader();
                int row = 2;
                while (sdr.Read())
                {
                    string pno = sdr[0].ToString();
                    string ad = sdr[1].ToString();
                    string soyad = sdr[2].ToString();
                    string semt = sdr[3].ToString();
                    string sehir = sdr[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + pno + "  " + ad + "  " + soyad + "  " + semt + "  " + sehir;
                    sayfa1.Cells[row, 1].Value2 = pno;
                    sayfa1.Cells[row, 2].Value2 = ad;
                    sayfa1.Cells[row, 3].Value2 = soyad;
                    sayfa1.Cells[row, 4].Value2 = semt;
                    sayfa1.Cells[row, 5].Value2 = sehir;
                    row++;
                }
                sdr.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred during Sql Query!!\n" + ex.ToString());
            }
            finally
            {
                if (baglanti.State == System.Data.ConnectionState.Open)
                    baglanti.Close();
            }
        }
        private void btnExceldenOku_Click(object sender, EventArgs e)
        {
            Excel.Application exlApp;
            Excel.Workbook exlWorkBook;
            Excel.Worksheet exlWorkSheet;
            Excel.Range range;
            int rowCnt = 0;
            int colCnt = 0;
            exlApp = new Excel.Application();
            exlWorkBook = exlApp.Workbooks.Open("C:\\Users\\Mustafa\\Desktop\\Personel.xlsx");
            exlWorkSheet = (Excel.Worksheet)exlWorkBook.Worksheets.get_Item(1);
            range = exlWorkSheet.UsedRange;
            richTextBox2.Clear();

            for (rowCnt = 2; rowCnt <= range.Rows.Count; rowCnt++)
            {
                ArrayList list = new ArrayList();

                for (colCnt = 1; colCnt <= range.Columns.Count; colCnt++)
                {
                    string okunanhucre = Convert.ToString((range.Cells[rowCnt, colCnt] as Excel.Range).Value2);
                    richTextBox2.Text = richTextBox2.Text + okunanhucre + "  ";
                    list.Add(okunanhucre);
                }
                richTextBox2.Text = richTextBox2.Text + "\n";
                try
                {
                    baglanti.Open();
                    SqlCommand sqlCommand = new SqlCommand("INSERT INTO Personel (PersonelNo, Ad, Soyad, Semt, Sehir) VALUES (@P1, @P2, @P3, @P4, @P5)", baglanti);
                    sqlCommand.Parameters.AddWithValue("@P1", list[0]);
                    sqlCommand.Parameters.AddWithValue("@P2", list[1]);
                    sqlCommand.Parameters.AddWithValue("@P3", list[2]);
                    sqlCommand.Parameters.AddWithValue("@P4", list[3]);
                    sqlCommand.Parameters.AddWithValue("@P5", list[4]);
                    sqlCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while writing to the database!!\n" + ex.ToString());
                }
                finally
                {
                    if (baglanti.State == System.Data.ConnectionState.Open)
                        baglanti.Close();
                }
            }
            exlApp.Quit();
            ReleaseObject(exlWorkSheet);
            ReleaseObject(exlWorkBook);
            ReleaseObject(exlApp);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
