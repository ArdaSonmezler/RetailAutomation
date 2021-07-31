using Spire.Xls;
using System.IO;
using System.Windows.Forms;
namespace $safeprojectname$
{
    class FileSettings
    {
        public string dosya = "dosyaYolu.txt";
        public string DosyaYolu = "";
        public string sheetName;
        Workbook workbook;
        Worksheet worksheet;
        StreamWriter sw;
        public string OpenFile()
        {
            OpenFileDialog openFile = new OpenFileDialog()
            {
                InitialDirectory = "C:",
                Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                Title = "Excel Dosyası Seçiniz..",
                RestoreDirectory = true,
                CheckFileExists = false
            };
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string[] lines = System.IO.File.ReadAllLines(dosya);
                DosyaYolu = openFile.FileName;
                sw = new StreamWriter(dosya);
                sw.WriteLine(DosyaYolu);
                sw.Close();
                workbook = new Workbook();
                workbook.LoadFromFile(lines[0]);
                worksheet = workbook.Worksheets[0];
                sheetName = worksheet.Name;
                string DosyaAdi = openFile.SafeFileName;
            }
            return DosyaYolu;
        }
    }
}
