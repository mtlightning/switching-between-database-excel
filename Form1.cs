using System.Data.SqlClient;
using System.Drawing.Text;
using Excel = Microsoft.Office.Interop.Excel; 

namespace ExcelVTEntegrasyonProjesi
{
    public partial class Form1 : Form
    {
        SqlConnection baglanti = new SqlConnection(@"Data Source=.;Initial Catalog=ProjelerVT;Integrated Security=True;Encrypt=False");



        public Form1()
        {
            InitializeComponent();
        }

        private void buttonVTdenOku_Click(object sender, EventArgs e)
        {
            Excel.Application excelUygulama = new Excel.Application();
            excelUygulama.Visible = true;
            Excel.Workbook workbook = excelUygulama.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sayfa1 = workbook.Sheets[1];
            string[] basliklar = { "Personel No", "Ad", "Soyad", "Semt", "Þehir" };
            Excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = sayfa1.Cells[1, (1 + i)];
                range.Value2 = basliklar[i];


            }



            try
            {
                baglanti.Open();
                string sqlCumlesi = "SELECT PersonelNO , Ad, Soyad, Semt , Sehir FROM Personel";
                SqlCommand sqlCommand = new SqlCommand(sqlCumlesi, baglanti);
                SqlDataReader reader = sqlCommand.ExecuteReader();


                int satir = 2; // ilk satýrýmýz(baþlýk hariç)


                while (reader.Read())
                {
                    string pno = reader[0].ToString();
                    string ad = reader[1].ToString();
                    string soyad = reader[2].ToString();
                    string semt = reader[3].ToString();
                    string sehir = reader[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + " " + pno + " " + ad + " " + soyad + " " + semt + " " + sehir + "\n";

                    range = sayfa1.Cells[satir, 1];
                    range.Value2 = pno;
                    range = sayfa1.Cells[satir, 2];
                    range.Value2 = ad;
                    range = sayfa1.Cells[satir, 3];
                    range.Value2 = soyad;
                    range = sayfa1.Cells[satir, 4];
                    range.Value2 = semt;
                    range = sayfa1.Cells[satir, 5];
                    range.Value2 = sehir;
                    satir++;

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL Query sýrasýnda bir hata oluþtu, HATA KODU : SQLREAD01 \n " + ex.ToString());


            }
            finally
            {
                if (baglanti != null)
                    baglanti.Close();
            }



        }

        private void buttonExceldenOku_Click(object sender, EventArgs e)
        {
            Excel.Application exlApp;
            Excel.Workbook exlWorkBook;
            Excel.Worksheet exlWorksheet;
            Excel.Range range;
            int rCnt = 0; ;
            int cCnt = 0;
            exlApp = new Excel.Application();
            exlWorkBook = exlApp.Workbooks.Open("C:\\testexcel\\Kitap1.xlsx"); 
            exlWorksheet = exlWorkBook.Worksheets.get_Item(1);
            range = exlWorksheet.UsedRange;

            // richTextBox2'nin içeriðini temizle
            richTextBox2.Clear();

            //ilk satýr baþlýklarý içerdiði için rCnt'ý 2'den baþlatmamýz gerekiyor.
            //ilk satýrda veriler baþlamýþ olsaydý 1'den baþlatmamýz gerekirdi.

            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                List<string> list = new List<string>();

                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    string okunanHucre = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    richTextBox2.Text = richTextBox2.Text + okunanHucre + "  ";
                    list.Add(okunanHucre);
                }

                richTextBox2.Text = richTextBox2.Text + "\n";

            }

            exlApp.Quit();
            ReleaseObject(exlWorksheet);
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
            catch(Exception ex)
            {
                obj  = null;

            }
            finally
            {
                GC.Collect();
            }
        }

        
    }
}
