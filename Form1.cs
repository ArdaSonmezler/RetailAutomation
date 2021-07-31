using BarcodeLib;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace $safeprojectname$
{
    public partial class Form1 : Form
    {
        FileSettings file = new FileSettings();
        public DataTable DTexcel;
        public Worksheet worksheet;
        public DataTable dt;
        public Form1()
        {
            InitializeComponent();
        }

        private void gunaBtnCreate_Click(object sender, EventArgs e)
        {
            //Barcode barcode2 = new Barcode();
            //int width = (int)(guna2PictureBox1.Width * 0.8);
            //int height = (int)(guna2PictureBox1.Height * 0.5);
            //Image image = barcode2.Encode(TYPE.CODE128, gunaTxtName.Text,Color.Black,Color.Transparent, width, height);
            Zen.Barcode.Code39BarcodeDraw barcode = Zen.Barcode.BarcodeDrawFactory.Code39WithChecksum;
            //gunaPicBox.Image = barcode.Draw(gunaTxtBarcode.Text, 40);
            gunaPicBox.Image = barcode.Draw(gunaTxtBarcode.Text,50);
            gunaPicBox.Paint += Guna2PictureBox_Paint_Price;
            gunaPicBox.Paint += Guna2PictureBox_Paint_Name;



            //dt.Rows.Clear();
            //double pFiyat = Convert.ToDouble(guna2TextBox5.Text);
            //double tFiyat = Convert.ToDouble(guna2TextBox4.Text);
            //dt.Rows.Add(guna2TextBox3.Text, guna2TextBox2.Text, pFiyat, tFiyat);
            //cry.Load(System.Windows.Forms.Application.StartupPath + "\\BARCODErpt2.rpt");
            //cry.SetDataSource(dt);

            //try
            //{
            //    ExportOptions CrExportOptions;
            //    DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            //    PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
            //    CrDiskFileDestinationOptions.DiskFileName = @"BARKOD.pdf";
            //    CrExportOptions = cry.ExportOptions;
            //    {
            //        CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
            //        CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
            //        CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
            //        CrExportOptions.FormatOptions = CrFormatTypeOptions;
            //    }
            //    cry.Export();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}

            //PdfLoadedDocument loadedDocument = new PdfLoadedDocument(@"BARKOD.pdf");
            //Bitmap image = loadedDocument.ExportAsImage(0);
            //image.Save(@"BARCODE.jpg", ImageFormat.Jpeg);
            //loadedDocument.Close(true);

            //Image img = Image.FromFile(@"BARCODE.jpg");
            //img.RotateFlip(RotateFlipType.Rotate90FlipNone);
            //img.Save(@"BARCODE.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            //pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            //pictureBox1.Image = img;
        }
        private void Guna2PictureBox_Paint_Name(object sender, PaintEventArgs e)
        {
            int w = Convert.ToInt32(gunaPicBox.Width * 0.1);
            int h = Convert.ToInt32(gunaPicBox.Height * 0.5);
            using (Font myFont = new Font("Arial", 15))
            {
                e.Graphics.DrawString(gunaTxtName.Text, myFont, Brushes.Black, new Point(w, h));
            }
        }
        private void Guna2PictureBox_Paint_Price(object sender, PaintEventArgs e)
        {
            int w = Convert.ToInt32(gunaPicBox.Width*0.1);
            int h = Convert.ToInt32(gunaPicBox.Height * 0.6);
            using (Font myFont = new Font("Arial", 15))
            {
                e.Graphics.DrawString(gunaTxtPrice.Text, myFont, Brushes.Black, new Point(w, h));
            }
        }
       
        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            try
            {
                string dosyaYolu = file.OpenFile();
                OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                baglanti.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + string.Concat(file.sheetName, "$") + "]", baglanti);
                DTexcel = new DataTable();
                da.Fill(DTexcel);
                gunaDataGridVİew.DataSource = DTexcel;
                baglanti.Close();
            }
            catch (OleDbException ex)
            {

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dt = new DataTable();
            dt.TableName = "URUNLER";
            dt.Columns.Add("BARKODNO", typeof(string));
            dt.Columns.Add("ÜRÜNADI", typeof(string));
            dt.Columns.Add("PESİN SATIŞ", typeof(double));
            dt.Columns.Add("TAKSİTLİ SATIŞ", typeof(double));
            if (File.Exists(file.dosya))
            {

                string[] lines = System.IO.File.ReadAllLines(file.dosya);
                if (File.Exists(lines[0]))
                {
                    Workbook workbook = new Workbook();
                    workbook.LoadFromFile(lines[0]);
                    worksheet = workbook.Worksheets[0];
                    string sheetName = worksheet.Name;
                    foreach (string line in lines)
                    {
                        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + line + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                        baglanti.Open();
                        OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "$]", baglanti);
                        DTexcel = new DataTable();
                        da.Fill(DTexcel);
                        gunaDataGridVİew.DataSource = DTexcel;
                        baglanti.Close();
                    }
                }

            }
            else
            {
                FileStream fs = File.Create(@file.dosya);

                fs.Close();
            }
        }

        private void gunaDataGridVİew_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            gunaTxtBarcode.Text = gunaDataGridVİew.CurrentRow.Cells[0].Value.ToString();
            gunaTxtName.Text = gunaDataGridVİew.CurrentRow.Cells[1].Value.ToString();
        }

        private void gunaTxtSearch_TextChanged(object sender, EventArgs e)
        {
            string cell = worksheet.Range["B1"].Value;
            DataView dv = DTexcel.DefaultView;
            dv.RowFilter = "[" + cell + "] LIKE '" + gunaTxtSearch.Text + "%'";
            gunaDataGridVİew.DataSource = dv;
        }

        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {
            PrintDialog pd = new PrintDialog();
            PrintDocument doc = new PrintDocument();
            doc.PrintPage += Doc_PrintPage;
            pd.Document = doc;
            if (pd.ShowDialog() == DialogResult.OK)
            {
                doc.Print();
            }
        }
        private void Doc_PrintPage(object sender, PrintPageEventArgs e)
        {
            Bitmap bm = new Bitmap(gunaPicBox.Width, gunaPicBox.Height);
            gunaPicBox.DrawToBitmap(bm, new Rectangle(0, 0, gunaPicBox.Width, gunaPicBox.Height));
            e.Graphics.DrawImage(bm, 0, 0);
            bm.Dispose();


        }
    }
}
