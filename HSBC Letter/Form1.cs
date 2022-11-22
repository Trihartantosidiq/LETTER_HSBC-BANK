using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf.codec;
using iTextSharp.text.pdf.qrcode;
using JohnLibs;
using System.Globalization;
using Aspose.Cells;

namespace HSBC_Letter
{
    public partial class Form1 : Form
    {
        Allmodul proc = new Allmodul();
        string dir_input = "";
        int jmldok = 0;
        CultureInfo ci = new CultureInfo("id-ID");
        CultureInfo di = new CultureInfo("id-EN");
        koneksi connect = new koneksi();
        string pathfont = Directory.GetCurrentDirectory() + "\\FONTS\\";

        public Form1()
        {
            InitializeComponent();
        }
        public BaseFont fonttable = BaseFont.CreateFont(Directory.GetCurrentDirectory() + @"\FONTS\ARIAL.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, true);
        private void dateTimePicker1_CloseUp(object sender, EventArgs e)
        {
            string tcycle = dateTimePicker1.Value.ToString("yyyyMMdd");
            textBox1.Text = tcycle;
            dir_input = Directory.GetCurrentDirectory() + @"\DATA\" + tcycle + @"\LOCK\";

            if (!Directory.Exists(dir_input))
            {
                MessageBox.Show("Tanggal Yang Anda Pilih Salah", "!Warning!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox2.Text = "";
                return;
            }
            else
            {
                string[] excel = Directory.GetFiles(dir_input, "*.xls*");
                
                textBox2.Text = excel[0];   
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            if (!this.backgroundWorker1.IsBusy)
            {
                this.backgroundWorker1.RunWorkerAsync();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int persen = (e.ProgressPercentage * 100) / jmldok;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Maximum = Convert.ToInt32(jmldok);
            progressBar1.Value = e.ProgressPercentage;
            label3.Text = Convert.ToString(persen) + "%";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                button1.Enabled = true;
                MessageBox.Show("Proses Letter HSBC WPB Network Finished \n Silakan Cek File Cetak & softcopynya", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {

                String tgcyc = dateTimePicker1.Value.ToString("yyyyMMdd");
                string fxls = textBox2.Text.Trim();
                string fdir = Path.GetDirectoryName(fxls);

                string filepassword = "";
                string namafileawal = "";
                string[] filep = Directory.GetFiles(fdir, "*.xlsx");
                foreach (string isi in filep)
                {
                    filepassword = isi;
                    namafileawal = Path.GetFileName(isi);
                }
                string filenpassword = Directory.GetCurrentDirectory() + "\\DATA\\" + tgcyc + "\\" + namafileawal;

                LoadOptions lo = new LoadOptions();
                //lo.Password = DateTime.Now.ToString("yyyyMMdd"); //ini klo berdasrakan cycle
                string Cyclepasswordexcelnya = "Hbidsdb1";
                lo.Password = Cyclepasswordexcelnya;

                Workbook workbook = new Workbook(filepassword, lo);
                workbook.Settings.Password = null;
                workbook.Save(filenpassword);



                
                string dircetak = @"CETAK\" + tgcyc + "\\";
                string dirsoft = @"SOFTCOPY\" + tgcyc + "\\";
                int seqdoc = 1;
                string docnoseq = seqdoc.ToString("D3");
                proc.buatdir(dircetak);
                proc.buatdir(dirsoft);
                string fhd = "NOSEQ;CYCLE;BRANCH SDB;NAMA BRANCH;COSTUMER NAME;ALAMAT1;ALAMAT2;ALAMAT3;ALAMAT4;JML_HAL;JML_AMPLOP;PRODUK;KURIR";
                string fsoft = dirsoft + "\\Softcopy-HSBC WPB Network-" + tgcyc + ".csv";
                string flcetak = dircetak + "\\" + docnoseq + "-HSBC WPB Network-" + tgcyc + ".PDF";
                using (System.IO.StreamWriter fs = new System.IO.StreamWriter(fsoft, false))
                {
                    fs.WriteLine(fhd);

                }
                DataTable dtxls = proc.excelapprove(filenpassword);
                int jmrec = dtxls.Rows.Count;
                jmldok = jmrec;
                int hitung = jmldok / 500 + 1;
                BaseFont FA = BaseFont.CreateFont(pathfont + "times.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, true);
                BaseFont FB = BaseFont.CreateFont(pathfont + "C39P36DlTt.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, true);
                BaseFont FC = BaseFont.CreateFont(pathfont + "timesbd.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, true);

                iTextSharp.text.Font FTimesOri = new iTextSharp.text.Font(FA, 10f, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font FTimerBolt = new iTextSharp.text.Font(FC, 10f, iTextSharp.text.Font.NORMAL);


                Paragraph Footer2 = new Paragraph();
                Footer2.Alignment = Element.ALIGN_JUSTIFIED;
                Footer2.SetLeading(0, 1.3f);


                int hitjmlpol = 0; int cekp = 0;
                string fldoc = Directory.GetCurrentDirectory() + "\\DATA\\" + tgcyc + "\\Template_HSBC_Letter.pdf";
                string tempbawah = Directory.GetCurrentDirectory() + "\\FORMS\\TemplateHSBC.pdf";
                string kopsurat = Directory.GetCurrentDirectory() + "\\FORMS\\kopsuratHSBC.pdf";

                if (!File.Exists(fldoc))
                { MessageBox.Show("Tidak ditemukan pdf :" + fldoc, "Info"); }
                PdfReader readletter = new PdfReader(fldoc);
                PdfReader readtempbawah = new PdfReader(tempbawah);
                PdfReader kopsuratHSBC = new PdfReader(kopsurat);

                int jmlpage = readletter.NumberOfPages;

                for (int qwe = 0; qwe < hitung; qwe++)
                {
                    if (hitjmlpol == 500)
                    {
                        seqdoc++;
                        docnoseq = seqdoc.ToString("D3");
                        hitjmlpol = 0;
                        flcetak = dircetak + "\\" + docnoseq + "-HSBC WPB Network-" + tgcyc + ".PDF";
                    }
                    Document doc = new Document(PageSize.A4);
                    PdfWriter writer = PdfAWriter.GetInstance(doc, new FileStream(flcetak, FileMode.Create));
                    writer.PdfVersion = PdfWriter.VERSION_1_5;
                    PdfContentByte pb;
                    doc.Open();
                    int nos = 0, thal = 0;
                    for (int rr = cekp; rr < jmrec; rr++)
                    {

                        if (hitjmlpol == 500)
                        {
                            break;
                        }
                        backgroundWorker1.ReportProgress(rr + 1);
                        float fleft = 71.5f, ftop = 729.50f;
                        nos++;

                        string branchsdb = dtxls.Rows[rr][0].ToString().Trim().Replace("'","`");
                        string branchname = dtxls.Rows[rr][1].ToString().Trim().Replace("'", "`");
                        string costumername = dtxls.Rows[rr][2].ToString().Trim().Replace("'", "`");
                        string add1 = dtxls.Rows[rr][3].ToString().Trim().Replace("'", "`");
                        string add2 = dtxls.Rows[rr][4].ToString().Trim().Replace("'", "`");
                        string add3 = dtxls.Rows[rr][5].ToString().Trim().Replace("'", "`");
                        string add4 = dtxls.Rows[rr][6].ToString().Trim().Replace("'", "`");
                        //DateTime mydate = DateTime.Parse(tgcyc);
                        //String PanggilDate = mydate.ToString("dd MMMM yyyy", ci);
                        //String PanggilDate2 = mydate.ToString("dd MMMM yyyy", di);

                        
                        string noseq = nos.ToString("D6");
                        string vbar = "HSDB"+ tgcyc + "-" + noseq;

                        string rowsoft = noseq + ";" + tgcyc + ";" + branchsdb + ";" + branchname + ";" + costumername + ";" + add1 + ";" + add2 + ";" + add3 + ";" + add4 + ";" + "2" + ";" + "1" + ";" + "HSBC WPB Netwok" + ";" + "NCS" + ";" + vbar + ";";
                        //string[] rowdb = { noseq, tgcyc, cnum, vbar, nama, vaddr1, vaddr2, vaddr3, vaddr4, vzip, "2", "1", "PAYDI_Pre-COMMS", "-", "Bapak/Ibu" };
                        string[] rowdb = { branchname, branchsdb, costumername, tgcyc, add1, add2, add3, add4, "2", "1", "HSBC WPB Netwok", "NCS", vbar};
                        //string[] rowdb1 = { noseq, vbar, nama, vaddr1, vaddr2, vaddr3, vaddr4, vzip, "2", "1", "PAYDI_Pre-COMMS", "-", "Bapak/Ibu" };
                        string[] rowdb1 = { branchname, branchsdb, tgcyc, add1, add2, add3, add4, "2", "1", "HSBC WPB Netwok", "NCS", vbar };
                        string kolom = "[Nama_Branch],[BranchSDB],[Cycle],[Add1],[Add2],[Add3],[Add4],[JML_HAL],[JML_AMPLOP],[PRODUK],[KURIR],[NOMOR_BARCODE]";
                        string[] kolomsplit = kolom.Split(',');
                        using (System.IO.StreamWriter fs = new System.IO.StreamWriter(fsoft, true))
                        {
                            string cek = "SELECT * FROM [HSBC_Letter].[dbo].[DATA] WHERE Cycle='" + tgcyc + "' AND [Costumer_Name]='" + costumername + "'";
                            DataTable detect = connect.openTable(cek);
                            //if (detect.Rows.Count > 0)
                            //{
                            //    cek = "UPDATE  [HSBC_Letter].[dbo].[DATA] SET ";
                            //    string[] set = kolomsplit.Zip(rowdb1, (a, b) => a + "='" + b + "'").ToArray();
                            //    cek += string.Join(", ", set) + " WHERE Cycle= '" + tgcyc + "' AND Costumer_Name= '" + costumername + "'";
                            //    Console.Write(cek);
                            //    connect.executeQuery(cek);
                            //}
                            //else
                            //{
                                string que = "INSERT INTO [HSBC_Letter].[dbo].[DATA] ([Nama_Branch],[BranchSDB],[Costumer_Name],[Cycle],[Add1],[Add2],[Add3],[Add4],[JML_HAL],[JML_AMPLOP],[PRODUK],[KURIR],[NOMOR_BARCODE])";
                                que += " VALUES('" + string.Join("', '", rowdb) + "')";
                                Console.Write(que);
                                connect.executeQuery(que);
                            //}
                            fs.WriteLine(rowsoft);
                            fs.Close();
                        }
                        hitjmlpol++; cekp++;

                        for (int P = 1; P <= jmlpage; P++)
                        {
                            bool lipat = false;
                            thal++;
                            pb = writer.DirectContent;
                            PdfImportedPage pageawal = writer.GetImportedPage(readletter, P);
                            PdfImportedPage kopHSBC = writer.GetImportedPage(kopsuratHSBC, 1);
                            doc.NewPage();
                            ColumnText ct = new ColumnText(pb);
                            string vfoota = tgcyc + "-" + noseq + "-" + P.ToString("D2") + "/" + jmlpage.ToString("D2");
                            pb.AddTemplate(pageawal, 0, 0);
                            pb.AddTemplate(kopHSBC, 0, 0);
                            pb.SetColorFill(BaseColor.BLACK);

                            if (P == 1)
                            {
                                pb.BeginText();
                                pb.SetFontAndSize(FC, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "HSDB/29/10/2022/" + noseq, fleft, ftop, 0);
                                ftop = ftop - 15;
                                pb.SetFontAndSize(FC, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Jakarta, " + "1 November 2022", fleft, ftop, 0);
                                ftop = ftop - 25;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Kepada ", fleft, ftop, 0);
                                ftop = ftop;
                                pb.SetFontAndSize(FC, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bapak/Ibu " + costumername, fleft + 34f, ftop, 0);
                                ftop = ftop - 12;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, add1, fleft, ftop, 0);
                                ftop = ftop - 12;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, add2, fleft, ftop, 0);
                                ftop = ftop - 12;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, add3, fleft, ftop, 0);
                                ftop = ftop - 12;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, add4, fleft, ftop, 0);
                                ftop = ftop - 18;
                                pb.SetFontAndSize(FB, 13);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*" + vbar + "*", fleft, ftop, 0);
                                ftop = ftop - 10;
                                pb.SetFontAndSize(FA, 7);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, vbar, fleft, ftop, 0);
                                pb.SetFontAndSize(FA, 9);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, vfoota, 520f, 15f, 0);


                                Chunk c1 = new Chunk("Sesuai dengan kebijakan Bank, bersama surat ini kami sampaikan bahwa efektif tanggal  ", FTimesOri);
                                Chunk c2 = new Chunk("1 Januari 2023", FTimerBolt);
                                Chunk c3 = new Chunk(", seluruh Kantor Cabang HSBC tidak lagi menyediakan fasilitas Kotak Penyimpanan/Safe Deposit Box (“SDB”). " +
                                    "Berdasarkan catatan kami Bapak/Ibu memiliki fasilitas SDB di HSBC Cabang ", FTimesOri);
                                Chunk c4 = new Chunk(branchname, FTimerBolt);
                                Chunk c5 = new Chunk(", sehingga kami mohon agar Bapak/Ibu dapat melakukan proses penutupan fasilitas SDB yang Bapak/Ibu miliki sebelum berakhirnya masa sewa SDB Bapak/Ibu. " +
                                    "Kami mohon maaf atas ketidaknyamanan yang mungkin Bapak/Ibu alami sehubungan dengan hal tersebut. ", FTimesOri);


                                Phrase p2 = new Phrase();

                                p2.Add(c1);
                                p2.Add(c2);
                                p2.Add(c3);
                                p2.Add(c4);
                                p2.Add(c5);
                                Footer2.Add(p2);
                                //Footer2.Font = FTimesOri;
                                //Footer2.Add("Melalui pemberitahuan ini, kami menginformasikan penyesuaian yang wajib Bapak/Ibu lakukan terhadap pilihan dana investasi pada polis " +
                                //    "Bapak / Ibu dengan Nomor Polis " + branchname);
                                ct.SetSimpleColumn(72f, 0, 523f, 510);
                                ct.AddElement(Footer2);
                                ct.Go();

                                Footer2.Clear();

                                PdfImportedPage tembawahid = writer.GetImportedPage(readtempbawah, 1);
                                pb.AddTemplate(tembawahid, 0, -10f);

                               
                                string strqr = P.ToString("D2") + jmlpage.ToString("D2") + noseq + "00000";
                                proc.ShowQRcode(strqr, doc, 85, 85, 0, 80, 45);
                                pb.EndText();
                            }
                            if (P == 2)
                            {
                                ftop = ftop + 112;
                                pb.BeginText();
                                pb.SetFontAndSize(FC, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "HSDB/29/10/2022/" + noseq, fleft, ftop, 0);
                                ftop = ftop - 15;
                                pb.SetFontAndSize(FC, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Jakarta, " + "1 November 2022", fleft, ftop, 0);
                                ftop = ftop - 25;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "To ", fleft, ftop, 0);
                                ftop = ftop;
                                pb.SetFontAndSize(FC, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Mr/Mrs " + costumername, fleft + 17f, ftop, 0);
                                ftop = ftop - 12;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, add1, fleft, ftop, 0);
                                ftop = ftop - 12;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, add2, fleft, ftop, 0);
                                ftop = ftop - 12;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, add3, fleft, ftop, 0);
                                ftop = ftop - 12;
                                pb.SetFontAndSize(FA, 10);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, add4, fleft, ftop, 0);
                                ftop = ftop - 18;
                                pb.SetFontAndSize(FB, 13);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*" + vbar + "*", fleft, ftop, 0);
                                ftop = ftop - 10;
                                pb.SetFontAndSize(FA, 7);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, vbar, fleft, ftop, 0);
                                pb.SetFontAndSize(FA, 9);
                                pb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, vfoota, 520f, 15f, 0);
                                pb.EndText();

                                Chunk c1 = new Chunk("In accordance to Bank policy, with this letter, we hereby to inform you that effective as of ", FTimesOri);
                                Chunk c2 = new Chunk("1 January 2023", FTimerBolt);
                                Chunk c3 = new Chunk(", all HSBC Branch Office will no longer provide Safe Deposit Box (”SDB”) facilities. " +
                                    "Based on our record you have SDB facility at HSBC branch ", FTimesOri);
                                Chunk c4 = new Chunk(branchname, FTimerBolt);
                                Chunk c5 = new Chunk(", hence we would like to seek your cooperation to close your SDB facility before the end of your SDB rental period. We apologize for any inconvenience you may encounter related to this matter. ", FTimesOri);


                                Phrase p = new Phrase();

                                p.Add(c1);
                                p.Add(c2);
                                p.Add(c3);
                                p.Add(c4);
                                p.Add(c5);
                                Footer2.Add(p);
                                //Footer2.Font = FTimesOri;
                                //Footer2.Add("Melalui pemberitahuan ini, kami menginformasikan penyesuaian yang wajib Bapak/Ibu lakukan terhadap pilihan dana investasi pada polis " +
                                //    "Bapak / Ibu dengan Nomor Polis " + branchname);
                                ct.SetSimpleColumn(72f, 0, 523f, 511);
                                ct.AddElement(Footer2);
                                ct.Go();

                                Footer2.Clear();

                                PdfImportedPage tembawahen = writer.GetImportedPage(readtempbawah, 2);
                                pb.AddTemplate(tembawahen, 0, +3);

                                string strqr = P.ToString("D2") + jmlpage.ToString("D2") + noseq + "00000";
                                proc.ShowQRcode(strqr, doc, 85, 85, 0, 80, 45);
                            }
                        }
                            
                        
                    }
                    doc.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void HSBC_LETTER_Load(object sender, EventArgs e)
        {
            string dbase = "SELECT * FROM [HSBC_Letter].[dbo].[DATA]";
            DataTable paydi = connect.openTable(dbase);
        }
        public void Export_ExcelReportApproval_Aspose(DataTable Tbl, string filename)
        {
            if (Tbl.Rows.Count == 0)
            {
                MessageBox.Show("Tidak ada data 1 pun");
            }
            else
            {
                new Aspose.Cells.License().SetLicense(LicenseHelper.License.LStream);
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear();

                Style styleHeader = workbook.CreateStyle();
                styleHeader.Font.Name = "Arial";
                styleHeader.Font.Size = 10;
                styleHeader.Font.IsBold = false;
                styleHeader.Pattern = BackgroundType.Solid;
                styleHeader.Font.Color = Color.Black;
                styleHeader.ForegroundColor = Color.Aqua;
                styleHeader.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
                styleHeader.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
                styleHeader.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
                styleHeader.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;

                Style styleData = workbook.CreateStyle();
                styleData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                styleData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                Style styleFooter = workbook.CreateStyle();
                styleFooter.Font.Name = "Arial";
                styleFooter.Font.Size = 10;
                styleFooter.Font.IsBold = false;
                styleFooter.Pattern = BackgroundType.Solid;
                styleFooter.Font.Color = Color.Black;
                styleFooter.ForegroundColor = Color.White;
                styleFooter.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleFooter.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleFooter.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                styleFooter.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                Style styletop = workbook.CreateStyle();
                styletop.Font.Name = "Arial";
                styletop.Font.Size = 12;
                styletop.Font.IsBold = true;
                Worksheet workSheet = workbook.Worksheets.Add("Sheet1");

                int k = 5;


                workSheet.Cells[0, 0].PutValue("APPROVAL Surat WPB Network HSBC");


                workSheet.Cells[2, 0].PutValue("CYCLE : " + dateTimePicker1.Text);
                workSheet.Pictures.Add(0, 3, Directory.GetCurrentDirectory() + "\\RDS.jpg", 120, 100);
                workSheet.Cells[10, 0].PutValue("Jakarta, ");
                workSheet.Cells[12, 0].PutValue("PREPARED BY, ");
                workSheet.Cells[12, 2].PutValue("APPROVED BY, ");
                workSheet.Cells[16, 0].PutValue("(                    )");
                workSheet.Cells[16, 2].PutValue("(                    )");

                int temp = 0;

                int[] lastrow = { 0, 0, 0, 0, 0 };


                for (int i = 0; i < Tbl.Columns.Count; i++)
                {
                    workSheet.Cells[0 + k, i].PutValue(Tbl.Columns[i].ColumnName.ToUpper());
                    workSheet.Cells[0 + k, i].SetStyle(styleHeader);
                }

                for (int i = 0; i < Tbl.Rows.Count; i++)
                {
                    int d3 = Convert.ToInt32(Tbl.Rows[i][2]);
                    temp = temp + d3;

                    for (int j = 0; j < Tbl.Columns.Count; j++)
                    {
                        workSheet.Cells[(i + 1 + k), j].PutValue(Tbl.Rows[i][j]);
                        workSheet.Cells[(i + 1 + k), j].SetStyle(styleData);

                        if (j >= 1)
                        {
                            lastrow[j] = lastrow[j] + Convert.ToInt32(Tbl.Rows[i][j]);
                            workSheet.Cells[0 + k, i].SetStyle(styleHeader);
                        }
                    }
                }

                workSheet.Cells[Tbl.Rows.Count + k + 1, 0].PutValue("");
                workSheet.Cells[Tbl.Rows.Count + k + 1, 0].SetStyle(styleData);
                workSheet.Cells[Tbl.Rows.Count + k + 1, 0].PutValue("Total");
                workSheet.Cells[Tbl.Rows.Count + k + 1, 1].SetStyle(styleData);

                for (int ea = 0; ea < 5; ea++)
                {
                    if (ea >= 1)
                    {
                        workSheet.Cells[Tbl.Rows.Count + k + 1, ea].PutValue(lastrow[ea]);
                        workSheet.Cells[Tbl.Rows.Count + k + 1, ea].SetStyle(styleData);
                    }
                }

                for (int x = 0; x < Tbl.Columns.Count; x++)
                {
                    workSheet.AutoFitColumn(x);
                }

                #region reg
                //for (int i = 0; i < Tbl.Columns.Count; i++)
                //{
                //    workSheet.Cells[0 + k, i].PutValue(Tbl.Columns[i].ColumnName);
                //    workSheet.Cells[0 + k, i].SetStyle(styleHeader);
                //}
                ////set datatabel in row
                //for (int i = 0; i < Tbl.Rows.Count; i++)
                //{
                //    for (int j = 0; j < Tbl.Columns.Count; j++)
                //    {

                //        workSheet.Cells[(i + 1 + k), j].PutValue(Tbl.Rows[i][j]);
                //        if ((i == Tbl.Rows.Count - 1) || (j == 0))
                //        {
                //            workSheet.Cells[(i + 1 + k), j].SetStyle(styleFooter);
                //        }
                //        else
                //        {
                //            workSheet.Cells[(i + 1 + k), j].SetStyle(styleData);
                //        }
                //    }
                //}
                //for (int x = 0; x < Tbl.Columns.Count; x++)
                //{
                //    workSheet.AutoFitColumn(x);
                //}
                #endregion

                workbook.Save(filename);
                Tbl.Clear();
                //MessageBox.Show("File Report Approval Billing OK");
            }
        }

        public void Export_ExcelLogKurir_Aspose(DataTable Tbl, string filename)
        {
            if (Tbl.Rows.Count == 0)
            {
                MessageBox.Show("Tidak ada data 1 pun");
            }
            else
            {
                new Aspose.Cells.License().SetLicense(LicenseHelper.License.LStream);
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear();

                Style styleHeader = workbook.CreateStyle();
                styleHeader.Font.Name = "Arial";
                styleHeader.Font.Size = 10;
                styleHeader.Font.IsBold = false;
                styleHeader.Pattern = BackgroundType.Solid;
                styleHeader.Font.Color = Color.Black;
                styleHeader.ForegroundColor = Color.Aqua;
                styleHeader.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
                styleHeader.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
                styleHeader.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
                styleHeader.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;

                Style styleData = workbook.CreateStyle();
                styleData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                styleData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                Style styleFooter = workbook.CreateStyle();
                styleFooter.Font.Name = "Arial";
                styleFooter.Font.Size = 10;
                styleFooter.Font.IsBold = false;
                styleFooter.Pattern = BackgroundType.Solid;
                styleFooter.Font.Color = Color.Black;
                styleFooter.ForegroundColor = Color.White;
                styleFooter.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleFooter.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleFooter.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                styleFooter.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                Style styletop = workbook.CreateStyle();
                styletop.Font.Name = "Arial";
                styletop.Font.Size = 12;
                styletop.Font.IsBold = true;
                Worksheet workSheet = workbook.Worksheets.Add("Sheet1");

                int k = 4;

                workSheet.Cells[2, 1].PutValue("REPORT LOG KURIR Surat Surat WPB Network HSBC " + dateTimePicker1.Text);

                for (int i = 0; i < Tbl.Columns.Count; i++)
                {
                    workSheet.Cells[0 + k, i].PutValue(Tbl.Columns[i].ColumnName);
                    workSheet.Cells[0 + k, i].SetStyle(styleHeader);
                }
                //set datatabel in row
                for (int i = 0; i < Tbl.Rows.Count; i++)
                {
                    for (int j = 0; j < Tbl.Columns.Count; j++)
                    {

                        workSheet.Cells[(i + 1 + k), j].PutValue(Tbl.Rows[i][j]);
                        if ((i == Tbl.Rows.Count - 1) || (j == 0))
                        {
                            workSheet.Cells[(i + 1 + k), j].SetStyle(styleFooter);
                        }
                        else
                        {
                            workSheet.Cells[(i + 1 + k), j].SetStyle(styleData);
                        }
                    }
                }
                for (int x = 0; x < Tbl.Columns.Count; x++)
                {
                    workSheet.AutoFitColumn(x);
                }
                workbook.Save(filename);
                //MessageBox.Show("File Log Billing Kurir NCS OK");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string cycle = dateTimePicker1.Value.ToString("yyyyMMdd");
                DataTable datatampung = new DataTable();
                datatampung.Columns.Add("Jenis Produk");
                datatampung.Columns.Add("Data Cetak");
                datatampung.Columns.Add("Jumlah Hal");
                datatampung.Columns.Add("NCS");
                datatampung.Columns.Add("Data Asli");

                string appr = "select(select distinct PRODUK from [DATA] where Cycle = '" + cycle + "') as [PRODUK], " +
                              "(SELECT COUNT(*) Costumer_Name from [DATA] where Cycle = '" + cycle+ "' and PRODUK = 'HSBC WPB Netwok') as [DATA CETAK], " +
                              "(SELECT sum(convert(float, JML_HAL)) as JML_HAL from [DATA] where Cycle = '" + cycle + "' and PRODUK = 'HSBC WPB Netwok') as [JML_HAL], " +
                              "(SELECT COUNT(*) KURIR from [DATA] where Cycle = '" + cycle + "' and KURIR = 'NCS' and PRODUK = 'HSBC WPB Netwok') as [NCS], " +
                              "(SELECT COUNT(*) Costumer_Name from [DATA] where Cycle = '" + cycle + "' and PRODUK = 'HSBC WPB Netwok') as [DATA ASLI]";

                DataTable utk_appr = connect.openTable(appr);
                string productSPKC = utk_appr.Rows[0][0].ToString();
                string datacetakSPKC = utk_appr.Rows[0][1].ToString();
                string halamanSPKC = utk_appr.Rows[0][2].ToString();
                string ncsSPKC = utk_appr.Rows[0][3].ToString();
                string dataasliSPKC = utk_appr.Rows[0][4].ToString();

                datatampung.Rows.Add(productSPKC, datacetakSPKC, halamanSPKC, ncsSPKC, dataasliSPKC);

                string txtout = Environment.CurrentDirectory + "\\Report\\" + cycle + "\\Report Approval WPB Network_" + cycle + ".xls";
                if (!Directory.Exists(Path.GetDirectoryName(txtout)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(txtout));
                }
                Export_ExcelReportApproval_Aspose(datatampung, txtout);

                MessageBox.Show("Report Approval WPB Network Done");

                string logr = "SELECT [Costumer_Name] as [REFERENCE] , Nama_Branch as [Nama_Branch], " +
                                   "LEFT(Add1, 10) as [Add1], [NOMOR_BARCODE] as [NOMOR_BARCODE] from [DATA] where " +
                                   "[PRODUK] = 'HSBC WPB Netwok' and Cycle = '" + cycle + "' and kurir = 'NCS';";
                DataTable DataLog = connect.openTable(logr);

                string txtoutpln = Environment.CurrentDirectory + "\\Report\\" + cycle + "\\Report Log Kurir WPB Network_" + cycle + ".xls";
                if (!Directory.Exists(Path.GetDirectoryName(txtoutpln)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(txtoutpln));
                }
                Export_ExcelLogKurir_Aspose(DataLog, txtoutpln);

            }
            catch (Exception ER)
            {

                MessageBox.Show(ER.ToString());
            }
        }
    }
}
