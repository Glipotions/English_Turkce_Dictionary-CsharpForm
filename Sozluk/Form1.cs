using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace Sozluk
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            ExcelDosyaOlustur();
            DosyaKontrolIslemleri();
            

        }
        private void ExcelDosyaOlustur()
        {
            string DosyaYolu = @"C:\English";
            if (Directory.Exists(DosyaYolu))
            {
                
            }
            else { Directory.CreateDirectory(@"C:\English"); }
            string DosyaYolu1 = @"C:\\English\\Kelime Çalışması.xls";
            if (System.IO.File.Exists(DosyaYolu1))
            {}
            else
            {
                Excel.Application xlOrn = new Microsoft.Office.Interop.Excel.Application();

                if (xlOrn == null)
                {
                    MessageBox.Show("Excel yüklü değil!!");
                    return;
                }

                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlOrn.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Cells[1, 1] = "Numara";
                xlWorkSheet.Cells[1, 2] = "Türkçe";
                xlWorkSheet.Cells[1, 3] = "English";
                xlWorkSheet.Cells[1, 4] = "Türkçe2";

                xlWorkBook.SaveAs("C:\\English\\Kelime Çalışması.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlOrn.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlOrn);

                MessageBox.Show("Excel dosyası komununda oluşturuldu!");
            }
            //string dosya_dizini = AppDomain.CurrentDomain.BaseDirectory.ToString() + "C:\\English\\Kelime Çalışması.xls";
            //if (File.Exists(dosya_dizini) == true)
            //{
            //}
        }
        private static void DosyaKontrolIslemleri()
        {
            
            if (Directory.Exists(@"C:\English"))
            { }
            else
            { Directory.CreateDirectory(@"C:\English"); }
         

        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\English\Kelime Çalışması.xls; Extended Properties='Excel 12.0 xml;'");
        //HDR=YES;


        private void button2_Click(object sender, EventArgs e)
        {
            
            TrToEng ff = new TrToEng();
            ff.Show();
            
        }


        private void button6_Click(object sender, EventArgs e)
        {
            
            EngToTr ff = new EngToTr();
            ff.Show();
            

        }

        private void button7_Click(object sender, EventArgs e)
        {
            
            kelimeEkle ff = new kelimeEkle();
            ff.Show();
        }
    }
}
