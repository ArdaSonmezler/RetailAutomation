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

namespace retailAutomation
{
    public partial class frmRetailAutomation : Form
    {
        FileSettings file = new FileSettings();
        public DataTable DTexcel;
        public Worksheet worksheet;
        public DataTable dt;
        public frmRetailAutomation()
        {
            InitializeComponent();
        }

        private void gunaBtnCreate_Click(object sender, EventArgs e)
        {
            //Zen.Barcode.Code39BarcodeDraw barcode = Zen.Barcode.BarcodeDrawFactory.Code39WithChecksum;
            //gunaPicBoxBarcode.Image = barcode.Draw(gunaTxtName.Text, 50);
            int w = Convert.ToInt32(gunatxtWidth.Text);
            int h = Convert.ToInt32(gunatxtHeight.Text);
            Size size = new Size(w, h);
            gunaPicBoxBarcode.Size = size;
            gunaPicBoxBarcode2.Size = size;
            gunaPicBoxBarcode.Visible = false;
            gunaPicBoxBarcode2.Visible = false;

            if (gunaRadioBtnHorizontalBarcode.Checked)
            {
                gunaPicBoxBarcode.Visible = true;
                gunaPicBoxBarcode2.Visible = false;
                Barcode barcode2 = new Barcode();
                int width = (int)(gunaPicBoxBarcode.Width * 0.8);
                int height = (int)(gunaPicBoxBarcode.Height * 0.5);
                Image image = barcode2.Encode(TYPE.CODE128, gunaTxtBarcode.Text, Color.Black, Color.Transparent, width, height);
                gunaPicBoxBarcode.Image = image;
                gunaPicBoxBarcode.Paint += GunaPicBoxBarcode_Paint_Name;
                gunaPicBoxBarcode.Paint += GunaPicBoxBarcode_Paint_Sale;
                gunaPicBoxBarcode.Paint += GunaPicBoxBarcode_Paint_Sale2;
            }
            else if(gunaRadioBtnVerticalBarcode.Checked)
            {
                gunaPicBoxBarcode2.Visible = true;
                gunaPicBoxBarcode.Visible = false;
                Barcode barcode2 = new Barcode();
                int width = (int)(gunaPicBoxBarcode2.Width * 1);
                int height = (int)(gunaPicBoxBarcode2.Height * 0.4);
                Image image = barcode2.Encode(TYPE.CODE128, gunaTxtBarcode.Text, Color.Black, Color.Transparent, width, height);
                image.RotateFlip(RotateFlipType.Rotate90FlipX);
                gunaPicBoxBarcode2.Image = image;
                gunaPicBoxBarcode2.Paint += GunaPicBoxBarcode_Paint_Name2;
                gunaPicBoxBarcode2.Paint += GunaPicBoxBarcode_Paint_Sale_2;
                gunaPicBoxBarcode2.Paint += GunaPicBoxBarcode_Paint_Sale2_2;
            }
            else
            {
                MessageBox.Show("SEÇİM YAPINIZ");
            }

            //.jpg Çıktısı Oluşturulacak.
        }

        private void GunaPicBoxBarcode_Paint_Name(object sender, PaintEventArgs e)
        {
            int w = Convert.ToInt32(gunaPicBoxBarcode.Width * 0.3);
            int h = Convert.ToInt32(gunaPicBoxBarcode.Height * 0.6);
            using (Font myFont = new Font("Arial", 10))
            {
                e.Graphics.DrawString(gunaTxtName.Text.ToUpper(), myFont, Brushes.Black, new Point(w, h));
            }
        }

        private void GunaPicBoxBarcode_Paint_Sale(object sender, PaintEventArgs e)
        {
            int w = Convert.ToInt32(gunaPicBoxBarcode.Width * 0.3);
            int h = Convert.ToInt32(gunaPicBoxBarcode.Height * 0.7);
            using (Font myFont = new Font("Arial", 10))
            {
                e.Graphics.DrawString("Peşin Satış:"+gunaTxtSale.Text, myFont, Brushes.Black, new Point(w, h));
            }
        }

        private void GunaPicBoxBarcode_Paint_Sale2(object sender, PaintEventArgs e)
        {
            int w = Convert.ToInt32(gunaPicBoxBarcode.Width * 0.3);
            int h = Convert.ToInt32(gunaPicBoxBarcode.Height * 0.8);
            using (Font myFont = new Font("Arial", 10))
            {
                e.Graphics.DrawString("Taksitli Satış:" + gunaTxtSale2.Text, myFont, Brushes.Black, new Point(w, h));
            }
        }

        //VERTICAL BARCODE METHODS
        private void GunaPicBoxBarcode_Paint_Name2(object sender, PaintEventArgs e)
        {
            int w = Convert.ToInt32(gunaPicBoxBarcode2.Width * 0.2);
            int h = Convert.ToInt32(gunaPicBoxBarcode2.Height * 0.5);
            using (Font myFont = new Font("Arial", 10))
            {
                e.Graphics.DrawString(gunaTxtName.Text.ToUpper(), myFont, Brushes.Black, new Point(w, h));
            }
        }
        private void GunaPicBoxBarcode_Paint_Sale_2(object sender, PaintEventArgs e)
        {
            int w = Convert.ToInt32(gunaPicBoxBarcode2.Width * 0.2);
            int h = Convert.ToInt32(gunaPicBoxBarcode2.Height * 0.6);
            using (Font myFont = new Font("Arial", 10))
            {
                e.Graphics.DrawString("Peşin Satış:" + gunaTxtSale.Text, myFont, Brushes.Black, new Point(w, h));
            }
        }
        private void GunaPicBoxBarcode_Paint_Sale2_2(object sender, PaintEventArgs e)
        {
            int w = Convert.ToInt32(gunaPicBoxBarcode2.Width*0.2);
            int h = Convert.ToInt32(gunaPicBoxBarcode2.Height * 0.7);
            using (Font myFont = new Font("Arial", 10))
            {
                e.Graphics.DrawString("Taksitli Satış:" + gunaTxtSale2.Text, myFont, Brushes.Black, new Point(w, h));
            }
        }

        private void guna2DataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            gunaTxtBarcode.Text = guna2DataGridView1.CurrentRow.Cells[0].Value.ToString();
            gunaTxtName.Text = guna2DataGridView1.CurrentRow.Cells[1].Value.ToString();
        }

        private void frmRetailAutomation_Load(object sender, EventArgs e)
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
                        guna2DataGridView1.DataSource = DTexcel;
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

        private void gunaBtnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                string dosyaYolu = file.OpenFile();
                OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                baglanti.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + string.Concat(file.sheetName, "$") + "]", baglanti);
                DTexcel = new DataTable();
                da.Fill(DTexcel);
                guna2DataGridView1.DataSource = DTexcel;
                baglanti.Close();
            }
            catch (OleDbException ex)
            {

            }
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
            string cell = worksheet.Range["B1"].Value;
            DataView dv = DTexcel.DefaultView;
            dv.RowFilter = "[" + cell + "] LIKE '" + guna2TextBox1.Text + "%'";
            guna2DataGridView1.DataSource = dv;
        }

        private void gunaBtnPrinter_Click(object sender, EventArgs e)
        {
            PrintDialog pd = new PrintDialog();
            PrintDocument doc = new PrintDocument();
            doc.PrintPage += Doc_PrintPage; //METHOD ÇALIŞMIYOR.
            pd.Document = doc;
            if (pd.ShowDialog() == DialogResult.OK)
            {
                doc.Print();
            }
        }
        private void Doc_PrintPage(object sender, PrintPageEventArgs e)
        {
            //2 TANE PİCTUREBOX VAR DÜZENLEME YAPILACAK!!!!!
            //Bitmap bm = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            //pictureBox1.DrawToBitmap(bm, new Rectangle(0, 0, pictureBox1.Width, pictureBox1.Height));
            //e.Graphics.DrawImage(bm, 0, 0);
            //bm.Dispose();


        }
    }
}
