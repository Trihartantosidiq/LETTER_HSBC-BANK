using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows.Forms;
using System.Management;
using System.Management.Instrumentation;
using System.Security;
using System.Security.Cryptography;
using iTextSharp.text.pdf.draw;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.codec;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf.qrcode;
using Aspose.Cells;

namespace HSBC_Letter
{
    class Allmodul
    {
        public void buatdir(string pathdir)
        {
            if (!Directory.Exists(pathdir))
            {
                Directory.CreateDirectory(pathdir);
            }
        }
        public void ShowQRcode(string textQR, Document doc, int lbr, int tinggi, int left, int top, int scale)
        {
            BarcodeQRCode qrcode = new BarcodeQRCode(textQR, lbr, tinggi, null);
            iTextSharp.text.Image qrImg = qrcode.GetImage();
            qrImg.SetAbsolutePosition(left, top);
            qrImg.ScalePercent(scale);
            doc.Add(qrImg);
        }
        public void StampQRcode(string textQR, PdfContentByte doc, int lbr, int tinggi, int left, int top, int scale)
        {
            BarcodeQRCode qrcode = new BarcodeQRCode(textQR, lbr, tinggi, null);
            iTextSharp.text.Image qrImg = qrcode.GetImage();
            qrImg.SetAbsolutePosition(left, top);
            qrImg.ScalePercent(scale);
            doc.AddImage(qrImg);
        }
        public void cetomr(PdfContentByte pbx, int ctr, bool lipat)
        {
            //buat OMR
            int[] omr = new int[16];
            for (int btg = 1; btg < 16; btg++)
            { omr[btg] = 0; }
            omr[1] = 1;
            if (lipat == true)
            { omr[2] = 1; omr[3] = 0; }
            else
            { omr[2] = 0; omr[3] = 1; }
            int count = ctr;
            int nseq = ctr % 7;
            if (nseq == 0)
            { nseq = 7; }
            for (int btg = 1; btg < 4; btg++)
            {
                double jm = nseq * (Math.Pow(2, -(btg - 1)));
                jm = Math.Floor(jm);
                double jm2 = Convert.ToInt32(jm) % 2;
                omr[btg + 3] = Convert.ToInt32(jm2);
            }
            double brs_omr = 175; int ckb = 0;
            for (int btg = 1; btg < 16; btg++)
            {
                if (btg > 3 && ckb == 0)
                { brs_omr = brs_omr - 1.5; ckb = 1; }
                brs_omr = brs_omr + 11;
                string vbrs_omr = brs_omr.ToString("###.##");
                if (omr[btg] == 1 || omr[btg] == 2)
                {
                    float baris = Convert.ToSingle(vbrs_omr);
                    pbx.SetColorFill(BaseColor.BLACK);
                    pbx.Rectangle(7, baris, 22, 2); //L,T,W,H   (x, y, lebarnya, tingginya) Tutup 
                    //pb.Rectangle(4, baris, 250, 140); 
                    pbx.Fill();
                }
            }
        }

        public DataTable excelapprove(string fullname)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + fullname + ";Extended Properties=Excel 8.0");
            DataTable ListXls = new DataTable();
            if (fullname.Contains("xlsx"))
            {
                connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullname + ";Extended Properties=Excel 12.0");
            }
            ListXls.Clear();
            connection.Open();
            DataTable Sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            connection.Close();
            //string[] dsdsds = Sheets.Rows[1][2].ToString();
            //foreach (DataRow dr in Sheets.Rows)
            //{
            //string sht = dr[2].ToString().Replace("'", "");
            string oooo = Sheets.Rows[1][2].ToString();
            string sht = Sheets.Rows[1][2].ToString().Replace("'", "");
            DataTable dt2 = new DataTable();
                DataSet ds2 = new DataSet();
                string sheet = sht; //.Replace("$","");
                var connectionString = "";
                connectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fullname);
                if (fullname.Contains("xlsx"))
                {
                    connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source={0}; Extended Properties=Excel 12.0;", fullname);
                }
                try
                {
                    var adapter2 = new OleDbDataAdapter("SELECT * FROM [" + sheet + "]", connectionString); //colon name
                    adapter2.Fill(ds2, sheet);
                    dt2 = ds2.Tables[sheet];
                    ListXls = dt2;
                    //break;
                }
                catch (Exception msg)
                {
                    ds2 = null;
                    dt2 = null;
                    MessageBox.Show("Error Baca sheet Excel :" + msg.ToString());
                }
            //}
            return ListXls;
        }

        public void exportexcel(DataTable Tbl, string filename, string namasheet)
        {
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            Style styles = workbook.CreateStyle();
            Style styleIndex = workbook.CreateStyle();
            Style styleHeader = workbook.CreateStyle();
            //int styleIndex = styles.Add();
            //Style styleHeader = styles[styleIndex];
            //styleIndex = styles.Add();
            styleHeader.Font.Name = "Tahoma";
            styleHeader.Font.Size = 9;
            styleHeader.Font.IsBold = true;
            styleHeader.HorizontalAlignment = TextAlignmentType.Center;
            styleHeader.VerticalAlignment = TextAlignmentType.Center;
            styleHeader.Pattern = BackgroundType.Solid;
            styleHeader.Font.Color = Color.Black;
            styleHeader.ForegroundColor = Color.LightGray;
            styleHeader.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            styleHeader.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            styleHeader.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            styleHeader.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;


            //styleIndex = styles.Add();
            //Style styleData = styles[styleIndex];
            Style styleData = workbook.CreateStyle();
            //Style styleData = styles(styleIndex);
            styleData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            styleData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            styleData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            styleData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            //styleIndex = styles.Add();
            //Style styleHeader2 = styles[styleIndex];
            Style styleHeader2 = workbook.CreateStyle();
            styleHeader2.Font.Name = "Arial";
            styleHeader2.Font.Size = 14;
            styleHeader2.Font.IsBold = true;

            Worksheet workSheet = workbook.Worksheets.Add(namasheet);
            int asd = 0;

            for (int i = 0; i < Tbl.Rows.Count; i++)
            {
                for (int j = 0; j < Tbl.Columns.Count; j++)
                {
                    if (i == 0)
                    {

                        workSheet.Cells[i, j].PutValue(Tbl.Columns[j].ColumnName);
                        workSheet.Cells[i, j].SetStyle(styleHeader);

                    }
                    workSheet.Cells[(i + 1 + asd), j].PutValue(Tbl.Rows[i][j].ToString());
                    workSheet.Cells[(i + 1 + asd), j].SetStyle(styleData);
                }
            }
            workbook.Save(filename);
        }
        public string showfield(string txtnil, int bts)
        {
            txtnil = txtnil.PadRight(bts);
            return txtnil;
        }
        public DataTable ReadExcelSheet(string xls)
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + xls + ";Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1'");
            OleDbCommand oleDbCmd = new OleDbCommand();
            con.Open();
            oleDbCmd.Connection = con;
            DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            DataTable dt2 = new DataTable();

            bool dapat = false;
            for (int a = 0; a < dt.Rows.Count; a++)
            {
                string firstExcelSheetName = dt.Rows[a][2].ToString().ToUpper();


                string query = "select * from [" + firstExcelSheetName + "]";
                OleDbDataAdapter data = new OleDbDataAdapter(query, con);
                data.TableMappings.Add("Table", "dtExcel");
                data.Fill(dt2);
                dapat = true;
            }
            return dt2;
        }
        public string[] SplitAddr(string txtaddr, int maxline, int brs)
        {
            string[] hasil = new string[brs];
            for (int x = 0; x < brs; x++)
            { hasil[x] = ""; }
            int pjg = txtaddr.Length;
            int pjx = 0;
            string upword = txtaddr.Trim();
            string mh = "";
            int loop = 0;
            while (loop < brs) // maximal baris 5 
            {
                if (pjg > maxline)
                {
                    for (int k = pjg - 1; k > 0; k--)
                    {
                        string mt = upword.Trim().Substring(k, 1);
                        if (mt == " " || mt == ":" || mt == "," || mt == ".")
                        {
                            if (pjg - mh.Length > maxline)
                            {
                                mh = mh + mt;
                            }
                            else
                            { break; }
                        }
                        else
                        { mh = mh + mt; }
                    }
                    pjx = mh.Length;
                    hasil[loop] = upword.Substring(0, pjg - pjx);
                    int prg = upword.Substring(0, pjg - pjx).Length;
                    upword = upword.Substring(prg);

                    pjg = upword.Trim().Length;
                    mh = "";
                    loop++;
                }
                else
                {
                    hasil[loop] = upword.Trim();
                    loop = 100;
                    break;
                }
            }
            return hasil;
        }
    }
}
