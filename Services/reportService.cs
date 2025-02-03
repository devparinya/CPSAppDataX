using CPSAppData.Models;
using CPSAppData.Services;
using CpsDataApp.Models;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using System.Collections;
using ZXing.Common;

namespace QueueAppManager.Service
{
    public class reportService
    {
        public string C2PathFile = string.Empty;
        public string TablePathFile = string.Empty;
        dateTimeHelper datehelper = new dateTimeHelper();
        string documentno = string.Empty;
        string workno = string.Empty;

        #region Print PDF
        public bool doCreateC2PDFReport(List<DataCPSCard> dataCPS, SettingData setdata)
        {
            GlobalFontSettings.FontResolver = new FileFontResolver();
            PdfDocument doc = new PdfDocument();
            double spaceH = 0.6;
            string lednumber = string.Empty;                    

            doc.Info.Title = "C2 CPS DATA";
            for (int i = 0; i < dataCPS.Count; i++)
            {
                lednumber = dataCPS[i].LedNumber??string.Empty;
                workno = dataCPS[0].WorkNo ?? "";
                documentno = string.Format("ที่ RC_บค {1}_{0}", string.IsNullOrEmpty(workno) ? lednumber: workno,setdata.FestNo);
                PdfPage page = doc.AddPage();
                page.Size = PdfSharp.PageSize.A4;
                page.Orientation = PdfSharp.PageOrientation.Portrait;                

                XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsUnit.Centimeter);
                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                XFont font = new XFont("thsarabun", 16, XFontStyleEx.Bold, options);
                XFont font_h_b = new XFont("thsarabun", 0.28, XFontStyleEx.Bold, options);
                XFont font_h = new XFont("thsarabun", 0.28, XFontStyleEx.Regular, options);
                XFont font_normalbold = new XFont("thsarabun", 0.493, XFontStyleEx.Bold, options);
                XFont font_normal = new XFont("thsarabun", 0.493, XFontStyleEx.Regular, options);
                XFont font_normal_small = new XFont("thsarabun", 0.35, XFontStyleEx.Regular, options);
                XFont font_normal_install_bold = new XFont("thsarabun", 0.39, XFontStyleEx.Bold, options);
                XFont font_normal_small_under = new XFont("thsarabun", 0.42, XFontStyleEx.Underline, options);
                XFont font_bold_small_under = new XFont("thsarabun", 0.42, XFontStyleEx.Bold | XFontStyleEx.Underline, options);
                XFont font_small_bold = new XFont("thsarabun", 0.35, XFontStyleEx.Bold, options);

                #region Header
                XImage image = XImage.FromFile(string.Format("{0}{1}", Application.StartupPath, "Images\\ktc_logo.png"));
                gfx.DrawImage(image, 2.65, 1.25,2.47, 1.8);
                gfx.DrawString("บริษัท บัตรกรุงไทย จำกัด (มหาชน)", font_h_b, XBrushes.Gray, new XRect(6.17, 1.23, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("591 อาคารสมัชชาวาณิช 2 ชั้น 14 ถนนสุขุมวิท แขวงคลองตันเหนือ เขตวัฒนา กรุงเทพฯ 10110", font_h, XBrushes.Gray, new XRect(6.17, 1.52, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("โทร: 02 123 5100 โทรสาร: 02 123 5190", font_h, XBrushes.Gray, new XRect(6.17, 1.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("ทะเบียนเลขที่ 0107545000110", font_h, XBrushes.Gray, new XRect(6.17, 2.08, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                gfx.DrawString("Krungthai Card Public Company Limited", font_h_b, XBrushes.Gray, new XRect(6.17, 2.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("591 United Business Center II, 14FL., Sukhumvit 33 Rd, North Klongton, Wattana, Bangkok 10110", font_h, XBrushes.Gray, new XRect(6.17, 2.75, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("Tel: 02 123 5100 Fax: 02 123 5190", font_h, XBrushes.Gray, new XRect(6.17, 3.03, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
      
                gfx.DrawString(documentno, font_normalbold, XBrushes.Black, new XRect(2.5, 3.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString(string.Format("วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.FestDate??"")), font_normalbold, XBrushes.Black, new XRect(-2.5, 4.63, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                #endregion
                #region Content 1
                gfx.DrawString("เรื่อง", font_normalbold, XBrushes.Black, new XRect(2.5, 5.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("ผลการพิจารณาข้อตกลงการประนอมหนี้ตามข้อเสนอของผู้ร้อง", font_normal, XBrushes.Black, new XRect(4.0, 5.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("อ้างถึง", font_normalbold, XBrushes.Black, new XRect(2.5, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                gfx.DrawString("บัญชีเคทีซี ของ", font_normal, XBrushes.Black, new XRect(4.0, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString(string.Format("{0}", dataCPS[i].CustomerName), font_normalbold, XBrushes.Black, new XRect(6.0, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                gfx.DrawString("หมายเลข", font_normal, XBrushes.Black, new XRect(4.0, 5.47 + spaceH + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString(string.Format("{0}", dataCPS[i].CardNo), font_normalbold, XBrushes.Black, new XRect(5.25, 5.47 + spaceH + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                string text1 = "ตามที่ บริษัท บัตรกรุงไทย จำกัด ( มหาชน ) หรือ “เคทีซี” ได้ยกเลิกสมาชิกพร้อมเรียกเก็บหนี้คืนทั้งหมด/";
                XRect rect1 = new XRect(4.0, 8.11, 14.5, 10);
                DrawJustifiedText(gfx, text1, font_normal, rect1);

                string text2 = "ได้ดำเนินคดีกับผู้ร้อง ซึ่งศาลได้มีคำพิพากษาถึงที่สุดแล้ว และ ผู้ร้องยื่นคำร้องขอไกล่เกลี่ยข้อพิพาทในชั้นบังคับคดี ในคดีของ";
                XRect rect2 = new XRect(2.5, 8.11 + spaceH, 16, 10);
                DrawJustifiedText(gfx, text2, font_normal, rect2);

                string courtname = string.Format("{0}", dataCPS[i].CourtName);
                gfx.DrawString(courtname, font_normalbold, XBrushes.Black, new XRect(2.5, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                string textredno = "หมายเลขคดีแดงที่";
                gfx.DrawString(textredno, font_normal, XBrushes.Black, new XRect(gfx.MeasureString(courtname, font_normalbold).Width + 2.8, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                double widthdata = gfx.MeasureString(textredno, font_normal).Width + gfx.MeasureString(courtname, font_normalbold).Width;
                string rednumber = string.Format("{0}", dataCPS[i].RedNo);
                gfx.DrawString(rednumber, font_normalbold, XBrushes.Black, new XRect(widthdata + 3, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                string textktc = "ระหว่าง บริษัท บัตรกรุงไทย จำกัด ( มหาชน ) โจทก์ กับ";
                gfx.DrawString(textktc, font_normal, XBrushes.Black, new XRect(-2.5, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);

                gfx.DrawString(string.Format("{0} จำเลย", dataCPS[i].CustomerName), font_normal, XBrushes.Black, new XRect(2.5, 8.9 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                string text3 = "ต่อมาผู้ร้องได้ขอลดหย่อนภาระหนี้ เพื่อชำระหนี้ให้กับบริษัทฯ เป็นการเสร็จสิ้นตามความสามารถการชำระหนี้";
                XRect rect3 = new XRect(4.0, 10.5, 14.5, 10);
                DrawJustifiedText(gfx, text3, font_normal, rect3);

                gfx.DrawString("ของผู้ร้องและบริษัทฯ ได้อนุมัติลดหย่อนภาระหนี้ โดยมีเงื่อนไขให้รับชำระหนี้เสร็จสิ้น ดังนี้", font_normal, XBrushes.Black, new XRect(2.5, 10.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                #endregion
                #region TABLE

                XBrush backgroundBrush = XBrushes.LightGray;
                XBrush whiteBrush = XBrushes.Gainsboro;
                //Table HD
                XPen pen = new XPen(XColors.Black, 0.01);
                gfx.DrawRectangle(pen, backgroundBrush, 2.2, 11.50, 4.28, 1.76);
                gfx.DrawRectangle(pen, backgroundBrush, 6.48, 11.50, 2.7, 1.76);
                gfx.DrawRectangle(pen, backgroundBrush, 9.18, 11.50, 2.9, 1.76);

                gfx.DrawRectangle(pen, backgroundBrush, 12.08, 11.50, 2.20, 1.13);
                gfx.DrawRectangle(pen, backgroundBrush, 14.28, 11.50, 2.30, 1.13);
                gfx.DrawRectangle(pen, backgroundBrush, 16.48, 11.50, 2.30, 1.13);

                gfx.DrawRectangle(pen, backgroundBrush, 12.08, 12.63, 2.20, 0.63);
                gfx.DrawRectangle(pen, backgroundBrush, 14.28, 12.63, 2.30, 0.63);
                gfx.DrawRectangle(pen, backgroundBrush, 16.48, 12.63, 2.30, 0.63);

                ////Detail

                gfx.DrawRectangle(pen, 2.2, 13.26, 4.28, 0.63);
                gfx.DrawRectangle(pen, 2.2, 13.89, 4.28, 0.63);
                gfx.DrawRectangle(pen, 2.2, 14.52, 4.28, 0.63);
                gfx.DrawRectangle(pen, 2.2, 15.15, 4.28, 0.63);
                gfx.DrawRectangle(pen, 2.2, 15.78, 4.28, 0.63);

                gfx.DrawRectangle(pen, 6.48, 13.26, 2.7, 0.63);
                gfx.DrawRectangle(pen, 6.48, 13.89, 2.7, 0.63);
                gfx.DrawRectangle(pen, 6.48, 14.52, 2.7, 0.63);
                gfx.DrawRectangle(pen, 6.48, 15.15, 2.7, 0.63);
                gfx.DrawRectangle(pen, 6.48, 15.78, 2.7, 0.63);

                gfx.DrawRectangle(pen, whiteBrush, 9.18, 13.26, 2.9, 0.63);
                gfx.DrawRectangle(pen, 9.18, 13.89, 2.9, 0.63);
                gfx.DrawRectangle(pen, 9.18, 14.52, 2.9, 0.63);
                gfx.DrawRectangle(pen, 9.18, 15.15, 2.9, 0.63);
                gfx.DrawRectangle(pen, 9.18, 15.78, 2.9, 0.63);


                gfx.DrawRectangle(pen, whiteBrush, 12.08, 13.26, 2.20, 0.63);
                gfx.DrawRectangle(pen, 12.08, 13.89, 2.20, 0.63);
                gfx.DrawRectangle(pen, 12.08, 14.52, 2.20, 0.63);
                gfx.DrawRectangle(pen, 12.08, 15.15, 2.20, 0.63);
                gfx.DrawRectangle(pen, 12.08, 15.78, 2.20, 0.63);

                gfx.DrawRectangle(pen, whiteBrush, 14.28, 13.26, 2.20, 0.63);
                gfx.DrawRectangle(pen, 14.28, 13.89, 2.20, 0.63);
                gfx.DrawRectangle(pen, 14.28, 14.52, 2.20, 0.63);
                gfx.DrawRectangle(pen, 14.28, 15.15, 2.20, 0.63);
                gfx.DrawRectangle(pen, 14.28, 15.78, 2.20, 0.63);

                gfx.DrawRectangle(pen, whiteBrush, 16.48, 13.26, 2.30, 0.63);
                gfx.DrawRectangle(pen, 16.48, 13.89, 2.30, 0.63);
                gfx.DrawRectangle(pen, 16.48, 14.52, 2.30, 0.63);
                gfx.DrawRectangle(pen, 16.48, 15.15, 2.30, 0.63);
                gfx.DrawRectangle(pen, 16.48, 15.78, 2.30, 0.63);

                gfx.DrawString("เงื่อนไขการชำระ", font_normal_install_bold, XBrushes.Black, new XRect(3.5, 12.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("จำนวนเงินที่ชำระ", font_normal_install_bold, XBrushes.Black, new XRect(6.85, 12.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("จำนวนเงินที่ชำระ", font_normal_install_bold, XBrushes.Black, new XRect(9.7, 12.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("ต่องวด", font_normal_install_bold, XBrushes.Black, new XRect(10.3, 12.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(12.85, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(12.4, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(12.45, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(15.05, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(14.6, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(14.65, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(17.25, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(16.8, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(16.85, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                double diameter = 0.35; // เส้นผ่านศูนย์กลางของวงกลมในเซนติเมตร
                gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 13.56 - diameter / 2, diameter, diameter);
                gfx.DrawString("ชำระปิดบัญชีคราวเดียว", font_normal_small, XBrushes.Black, new XRect(2.85, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(10.6, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(13.1, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(15.3, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(17.6, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 14.19 - diameter / 2, diameter, diameter);
                gfx.DrawString("ผ่อนชำระไม่เกิน          6          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 13.98, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 14.82 - diameter / 2, diameter, diameter);
                gfx.DrawString("ผ่อนชำระไม่เกิน        12          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 14.61, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 15.45 - diameter / 2, diameter, diameter);
                gfx.DrawString("ผ่อนชำระไม่เกิน        24          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 15.24, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 16.08 - diameter / 2, diameter, diameter);
                gfx.DrawString("ผ่อนชำระไม่เกิน ......................... งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 15.87, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);



                #endregion
                #region Content 2
                gfx.DrawString("โดยเริ่มชำระงวดแรกภายในวันที่ .................................. งวดต่อไปชำระ ทุกๆ วันที่ ...................... ของทุกเดือน โดย", font_normal, XBrushes.Black, new XRect(4.0, 16.65, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("ชำระให้เสร็จสิ้นภายในวันที่ ...............................................ทั้งนี้บริษัทฯ อาจจะมีการแจ้งเตือนก่อนวันถึงกำหนดชำระผ่าน SMS", font_normal, XBrushes.Black, new XRect(2.5, 16.65+spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                string custtel = string.Format("{0}", dataCPS[i].CustomerTel);
                string textcusttel = "ตามหมายเลขโทรศัพท์ที่ท่านได้ให้ไว้แก่บริษัทฯ ครั้งหลังสุด";
                string textcusttelother = "หรือ ...................................................";
                gfx.DrawString(textcusttel, font_normal, XBrushes.Black, new XRect(2.5, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString(custtel, font_normalbold, XBrushes.Black, new XRect(gfx.MeasureString(textcusttel, font_normal).Width + 2.65, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                double withcustel = gfx.MeasureString(textcusttel, font_normal).Width + gfx.MeasureString(custtel, font_normalbold).Width;
                gfx.DrawString(textcusttelother, font_normal, XBrushes.Black, new XRect(withcustel + 2.8, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                string text4 = "ภายในกำหนดข้างต้นหากผู้ร้องชำระหนี้ดังกล่าวให้กับบริษัทฯเป็นที่ครบถ้วนเรียบร้อยแล้ว บริษัทฯจะถือการชำระ";
                XRect rect4 = new XRect(4.0, 18.90, 14.5, 10);
                DrawJustifiedText(gfx, text4, font_normal, rect4);

                string text5 = "หนี้ตามจำนวนดังกล่าวเป็นการชำระหนี้เสร็จสิ้นตามที่บริษัทฯอนุมัติ โดยบริษัทฯจักได้ดำเนินการปรับปรุงบัญชีของผู้ร้องต่อไป";
                XRect rect5 = new XRect(2.5, 18.90+spaceH, 16, 10);
                DrawJustifiedText(gfx, text5, font_normal, rect5);

                string txtcontact = "ติดต่อพนักงานผู้รับผิดชอบ ";
                string txtcollecttor = string.Format("{0} ", dataCPS[i].CollectorName);
                string txttel = "โทร. ";
                string txtcolltel = string.Format("{0}", dataCPS[i].CollectorTel);
                gfx.DrawString(txtcontact, font_normal_small_under, XBrushes.Black, new XRect(2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString(txtcollecttor, font_bold_small_under, XBrushes.Black, new XRect(gfx.MeasureString(txtcontact, font_normal_small_under).Width + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                double withcontact = gfx.MeasureString(txtcontact, font_normal_small_under).Width + gfx.MeasureString(txtcollecttor, font_bold_small_under).Width;
                gfx.DrawString(txttel, font_normal_small_under, XBrushes.Black, new XRect(withcontact + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString(txtcolltel, font_bold_small_under, XBrushes.Black, new XRect(withcontact + gfx.MeasureString(txttel, font_normal_small_under).Width+2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                gfx.DrawString("หมายเหตุ", font_normal_small, XBrushes.Black, new XRect(2.5,21.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("1. กรณีมีการชำระหนี้ข้างต้นถือเป็นการชำระหนี้เสร็จสิ้นเฉพาะบัตรเลขที่ดังกล่าวข้างต้นเท่านั้น ไม่รวมถึงภาระหนี้อื่นของ", font_normal_small, XBrushes.Black, new XRect(4.0, 21.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("บมจ.บัตรกรุงไทย (ถ้ามี)", font_normal_small, XBrushes.Black, new XRect(4.0, 21.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("2. หากผู้ร้องผิดนัด ไม่ชำระหนี้ตามข้อตกลงไม่ว่าข้อหนึ่งข้อใด ถือว่าข้อตกลงไหล่เกลี่ยเป็นอันยกเลิก และบมจ.บัตรกรุงไทยจะดำเนินการ", font_normal_small, XBrushes.Black, new XRect(4.0, 22.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("บังคับคดีตามกฎหมายต่อไป โดยกลับไปคิดยอดหนี้ตามคำพิพากษาของศาลส่วนเงินที่ผู้ร้องชำระเข้ามาตามข้อตกลงไกล่เกลี่ย ให้เป็นส่วน", font_normal_small, XBrushes.Black, new XRect(4.0, 22.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("หนึ่งของการชำระหนี้ตามคำพิพากษา (ถ้ามี)", font_normal_small, XBrushes.Black, new XRect(4.0, 23.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                gfx.DrawString("ลงชื่อ .......................................................... ผู้ร้องขอไกล่เกลี่ย", font_normal, XBrushes.Black, new XRect(2.5,25.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("ลงชื่อ .......................................................... ผู้แทนโจทก์", font_normal, XBrushes.Black, new XRect(-2.5, 25.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                gfx.DrawString("(                                              )", font_normal, XBrushes.Black, new XRect(3.1, 25.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                gfx.DrawString("(                                              )", font_normal, XBrushes.Black, new XRect(-3.85, 25.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                gfx.DrawString(string.Format("ลำดับกรม {0}", dataCPS[i].LedNumber), font_small_bold, XBrushes.Black, new XRect(-2.5, 27.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                #endregion

            }
            if (dataCPS.Count > 0)
            {
                string file_name = string.Format("C2_{0}.pdf", dataCPS[0].CustomerID);
                file_name = file_name.Replace("\n", "").Replace("\r", "").Replace("/", "").Replace(" ", "");
                string fullfilename = Path.Combine(C2PathFile, file_name);
                if (Path.Exists(C2PathFile))
                {
                    doc.Save(fullfilename);
                    return true;
                }               
            }
            return false;
        }
        public bool doCreateCPSTableReport(List<DataCPSCard> dataCPS,SettingData setdata)
        {
            GlobalFontSettings.FontResolver = new FileFontResolver();
            PdfDocument doc = new PdfDocument();

            double sumjudgmentAmnt = 0;
            double sumcapitalAmnt = 0;
            double sumdeptAmnt = 0;
            double sumaccCloseAmnt = 0;
            double sumaccClose6Amnt = 0;
            double suminstallment6Amnt = 0;
            double sumaccClose12Amnt = 0;
            double suminstallment12Amnt = 0;
            double sumaccClose24Amnt = 0;
            double suminstallment24Amnt = 0;

            doc.Info.Title = "CPS Table";
            PdfPage page = doc.AddPage();
            page.Size = PdfSharp.PageSize.A4;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsUnit.Centimeter);
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            XFont font_title_bold = new XFont("thsarabun", 0.56, XFontStyleEx.Bold, options);
            XFont font_normalbold = new XFont("thsarabun", 0.49, XFontStyleEx.Bold, options);
            XFont font_normal = new XFont("thsarabun", 0.49, XFontStyleEx.Regular, options);
            XFont font_normal_underline = new XFont("thsarabun", 0.49, XFontStyleEx.Underline, options);
            XFont font_normal_small = new XFont("thsarabun", 0.38, XFontStyleEx.Regular, options);
            XFont font_bold_small = new XFont("thsarabun", 0.38, XFontStyleEx.Bold, options);
            XPen pen = new XPen(XColors.Black, 0.03);

            XImage image = XImage.FromFile(string.Format("{0}{1}", Application.StartupPath, @"Images/ktc_logo.png"));
            gfx.DrawImage(image, 1.1, 1.72, 3.24, 2.33);

            if (!string.IsNullOrEmpty(dataCPS[0].LedNumber))
            {
                BitMatrix Qrcode = qrcodeService.Encode(dataCPS[0].LedNumber ?? "",1);
                qrcodeService.DrawQrCode(gfx, new XPoint(25.6, 1.3), Qrcode,0.10);
            }
            gfx.DrawString(string.Format("มหกรรมไกล่เกลี่ยชั้นบังคับคดี ครั้งที่ {0}",setdata.FestNo), font_title_bold, XBrushes.Black, new XRect(0, 1.72, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("{0}",setdata.FestName), font_title_bold, XBrushes.Black, new XRect(0, 2.4, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.FestDate??"")), font_title_bold, XBrushes.Black, new XRect(0, 3.1, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

            gfx.DrawString("ลำดับที่", font_normalbold, XBrushes.Black, new XRect(1.2, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].WorkNo)?"-":dataCPS[0].WorkNo), font_normal, XBrushes.Black, new XRect(2.3, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
           
            gfx.DrawString("ลำดับกรม", font_normalbold, XBrushes.Black, new XRect(-1, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].LedNumber)?"-": dataCPS[0].LedNumber), font_normal, XBrushes.Black, new XRect(0.2, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

            gfx.DrawString("เลขที่คดีดำ", font_normalbold, XBrushes.Black, new XRect(21.5, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].BlackNo) ? "-" : dataCPS[0].BlackNo), font_normal, XBrushes.Black, new XRect(23.1, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ชื่อ", font_normalbold, XBrushes.Black, new XRect(1.2, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].CustomerName), font_normal, XBrushes.Black, new XRect(1.7, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString(string.Format("ภาระหนี้ ณ วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.DateAtCalulate??"")), font_normalbold, XBrushes.Black, new XRect(-0.3, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
           
            gfx.DrawString("เลขที่คดีแดง", font_normalbold, XBrushes.Black, new XRect(21.5, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].RedNo), font_normal, XBrushes.Black, new XRect(23.25, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("วันพิพากษา", font_normalbold, XBrushes.Black, new XRect(21.5, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", datehelper.doGetShortDateTHFromDBToPDF(dataCPS[0].JudgeDate ?? "")), font_normal, XBrushes.Black, new XRect(23.2, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            double diameter = 0.48; // เส้นผ่านศูนย์กลางของวงกลมในเซนติเมตร
            gfx.DrawEllipse(pen, XBrushes.White, 1.6 - diameter / 2, 6.75 - diameter / 2, diameter, diameter);
            gfx.DrawString("ตกลงรับเงื่อนไข", font_normal, XBrushes.Black, new XRect(2.1, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawEllipse(pen, XBrushes.White, 5.6 - diameter / 2, 6.75 - diameter / 2, diameter, diameter);
            gfx.DrawString("ไม่ตกลงรับเงื่อนไข  เนื่องจาก ............................................................", font_normal, XBrushes.Black, new XRect(6.1, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);           

            #region TABLE
            #region HEADER
            XPen pentable = new XPen(XColors.Black, 0.015);
            gfx.DrawRectangle(pentable, 1.29, 7.4, 0.85, 2.1);
            gfx.DrawRectangle(pentable, 2.14, 7.4, 3.58, 2.1);
            gfx.DrawRectangle(pentable, 5.72, 7.4, 2.61, 2.1);
            gfx.DrawRectangle(pentable, 8.33, 7.4, 2.03, 2.1);
            gfx.DrawRectangle(pentable, 10.36,7.4, 2.47, 2.1);
            gfx.DrawRectangle(pentable, 12.83, 7.4, 15.72, 2.1);
            gfx.DrawString("ที่", font_normal, XBrushes.Black, new XRect(1.62, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("หมายเลขบัตร", font_normal, XBrushes.Black, new XRect(3.0, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("พิพากษาทั้งสิ้น", font_normal, XBrushes.Black, new XRect(6.1, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("เงินต้น", font_normal, XBrushes.Black, new XRect(8.9, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ภาระหนี้", font_normal, XBrushes.Black, new XRect(11.0, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ปัจจุบัน", font_normal, XBrushes.Black, new XRect(11.1, 8.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("เงื่อนไขในการชำระ", font_normal, XBrushes.Black, new XRect(19.4, 7.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            

            gfx.DrawRectangle(pentable, 12.83, 8.1, 2.7, 0.7);
            gfx.DrawRectangle(pentable, 12.83, 8.8, 2.7, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 13.3 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ปิดงวดเดียว", font_normal, XBrushes.Black, new XRect(13.75, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(13.65, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 15.53, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 15.53, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 17.70, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 16.0 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 6 งวด", font_normal, XBrushes.Black, new XRect(17.2, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(16.10, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(18.4, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 19.87, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 19.87, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 22.04, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 20.4 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 12 งวด", font_normal, XBrushes.Black, new XRect(21.5, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(20.5, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(22.8, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 24.21, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 24.21, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 26.38, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 24.8 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 24 งวด", font_normal, XBrushes.Black, new XRect(25.9, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(24.9, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(27.22, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            #endregion
            #region Detail
            double ypos_strat = 9.5;
            double ypostext_strat = 9.55;
            string dat = "-";

            for (int i = 0; i < 7; i++)
            {
                if (i != 6) 
                { 
                    gfx.DrawRectangle(pentable, 1.29, ypos_strat, 0.85, 0.7);
                    gfx.DrawString(string.Format("{0}", i + 1), font_normal, XBrushes.Black, new XRect(1.62, ypos_strat+0.05, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);                   
                }              
                gfx.DrawRectangle(pentable, 2.14, ypos_strat, 3.58, 0.7);
                gfx.DrawRectangle(pentable, 5.72, ypos_strat, 2.61, 0.7);
                gfx.DrawRectangle(pentable, 8.33, ypos_strat, 2.03, 0.7);
                gfx.DrawRectangle(pentable, 10.36, ypos_strat,2.47, 0.7);
                gfx.DrawRectangle(pentable, 12.83, ypos_strat,2.7, 0.7);
                gfx.DrawRectangle(pentable, 15.53, ypos_strat,2.17, 0.7);
                gfx.DrawRectangle(pentable, 17.70, ypos_strat,2.17, 0.7);
                gfx.DrawRectangle(pentable, 19.87, ypos_strat,2.17,0.7);
                gfx.DrawRectangle(pentable, 22.04, ypos_strat,2.17,0.7);
                gfx.DrawRectangle(pentable, 24.21, ypos_strat,2.17,0.7);
                gfx.DrawRectangle(pentable, 26.38, ypos_strat,2.17,0.7);

                if (i < dataCPS.Count) // SET VALE FROM DB
                {
                    #region calculate value
                    double judgmentAmnt =  dataCPS[i].JudgmentAmnt;
                    double capitalAmnt = dataCPS[i].PrincipleAmnt; //CapitalAmnt;
                    double deptAmnt = dataCPS[i].DeptAmnt;
                    double accCloseAmnt = dataCPS[i].AccCloseAmnt;
                    double accClose6Amnt = dataCPS[i].AccClose6Amnt;
                    double installment6Amnt = dataCPS[i].Installment6Amnt;
                    double accClose12Amnt = dataCPS[i].AccClose12Amnt;
                    double installment12Amnt = dataCPS[i].Installment12Amnt;
                    double accClose24Amnt = dataCPS[i].AccClose24Amnt;
                    double installment24Amnt = dataCPS[i].Installment24Amnt;
                    string cardno = string.Format("{0}", dataCPS[i].CardNo);

                    if (dataCPS[i].Maxmonth < 6)
                    {
                        accClose6Amnt = 0;
                        installment6Amnt = 0;
                    }
                    if (dataCPS[i].Maxmonth < 12)
                    {
                        accClose12Amnt = 0;
                        installment12Amnt = 0;
                    }
                    if (dataCPS[i].Maxmonth < 24)
                    {
                        accClose24Amnt = 0;
                        installment24Amnt = 0;
                    }

                    sumjudgmentAmnt = sumjudgmentAmnt + judgmentAmnt;
                    sumcapitalAmnt = sumcapitalAmnt + capitalAmnt;
                    sumdeptAmnt = sumdeptAmnt + deptAmnt;
                    sumaccCloseAmnt = sumaccCloseAmnt + accCloseAmnt;
                    sumaccClose6Amnt = sumaccClose6Amnt + accClose6Amnt;
                    suminstallment6Amnt = suminstallment6Amnt + installment6Amnt;
                    sumaccClose12Amnt = sumaccClose12Amnt + accClose12Amnt;
                    suminstallment12Amnt = suminstallment12Amnt + installment12Amnt;
                    sumaccClose24Amnt = sumaccClose24Amnt + accClose24Amnt;
                    suminstallment24Amnt = suminstallment24Amnt + installment24Amnt;
                    #endregion

                    #region SET Value
                    if (!string.IsNullOrEmpty(cardno))
                    {
                        gfx.DrawString(cardno, font_normal, XBrushes.Black, new XRect(2.4, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(3.9, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (judgmentAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", judgmentAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-21.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (capitalAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", capitalAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-19.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (deptAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", deptAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-17.0, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accCloseAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accCloseAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-14.3, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose6Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-12.1, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment6Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-9.95, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose12Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-7.80, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment12Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-5.58, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose24Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-3.4, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment24Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-1.25, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    #endregion

                }
                else
                {
                    #region SET SUM VALUE
                    if (i == 6)
                    {
                        gfx.DrawString("รวม", font_normal_underline, XBrushes.Black, new XRect(3.69, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        if (sumjudgmentAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumjudgmentAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-21.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumcapitalAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumcapitalAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-19.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumdeptAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumdeptAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-17.0, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccCloseAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccCloseAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-14.3, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccClose6Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-12.1, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (suminstallment6Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-9.95, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccClose12Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-7.80, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (suminstallment12Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-5.58, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (sumaccClose24Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-3.4, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (suminstallment24Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-1.25, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(3.9, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    #endregion                   
                }               
                ypostext_strat = ypostext_strat + 0.7;
                ypos_strat = ypos_strat + 0.7;                              
            }
            #region TP LP
            double y_point = 15.2;
            List<string> datashow = CreateTPDP(dataCPS);
            for (int n = 0; n < 6; n++)
            {
                gfx.DrawString(datashow[n], font_normal_small, XBrushes.Black, new XRect(12, y_point, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                y_point = y_point + 0.7;
            }
            #endregion
            #endregion
            #endregion
            gfx.DrawString("สถานะทางคดี :", font_bold_small, XBrushes.Black, new XRect(2.14,15.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].LegalStatus), font_normal_small, XBrushes.Black, new XRect(3.8, 15.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("หมายเหตุบังคับคดี :", font_bold_small, XBrushes.Black, new XRect(2.14, 15.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].LegalExecRemark) ? "-": dataCPS[0].LegalExecRemark), font_normal_small, XBrushes.Black, new XRect(4.3, 15.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("วันที่ยึดทรัพย์/อายัดเงินเดือน :", font_bold_small, XBrushes.Black, new XRect(2.14, 16.6, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}",datehelper.doGetDateThaiFromDBToPDF(dataCPS[0].LegalExecDate ?? "-")), font_normal_small, XBrushes.Black, new XRect(5.4, 16.6, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ตรวจสอบ : _________________________________", font_normalbold, XBrushes.Black, new XRect(2.14, 18.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ผู้เจรจา : __________________________________", font_normalbold, XBrushes.Black, new XRect(18.4, 16.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString(string.Format("เจ้าหน้าที่ : | {0} | {1} | เบอร์ติดต่อ : {2}", dataCPS[0].CollectorName, string.IsNullOrEmpty(dataCPS[0].CollectorTeam) ? "-": dataCPS[0].CollectorTeam, dataCPS[0].CollectorTel), font_normalbold, XBrushes.Black, new XRect(-3.5, 18.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);

            if (dataCPS.Count > 0)
            {
                string file_name = string.Format("Table_{0}.pdf", dataCPS[0].CustomerID);
                file_name = file_name.Replace("\n", "").Replace("\r", "").Replace("/", "").Replace(" ", "");
                string fullfilename = Path.Combine(TablePathFile, file_name);
                if (Path.Exists(TablePathFile))
                {
                    doc.Save(fullfilename);
                    return true;
                }
            }
            return false;
        }
        public bool doCreateC2PDFMerge(ArrayList cpscardlist,SettingData setdata,string typerange) 
        {
            GlobalFontSettings.FontResolver = new FileFontResolver();
            PdfDocument doc = new PdfDocument();
            List<DataCPSCard> dataCPS = new List<DataCPSCard>();
            double spaceH = 0.6;
            string lednumber = string.Empty;
            string f_lednumber = string.Empty; 
            string l_lednumber = string.Empty;
            string f_workno = string.Empty;
            string l_workno = string.Empty;
            for (int n = 0; n < cpscardlist.Count; n++)
            {
                var cpscard = cpscardlist[n];
                
                if (cpscard != null)                
                {
                    dataCPS = (List<DataCPSCard>)cpscard;
                    if (n == 0) f_lednumber = dataCPS[0].LedNumber ?? string.Empty;
                    if (n == cpscardlist.Count-1) l_lednumber = dataCPS[0].LedNumber ?? string.Empty;
                    if (n == 0) f_workno = dataCPS[0].WorkNo ?? string.Empty;
                    if (n == cpscardlist.Count - 1) l_workno = dataCPS[0].WorkNo ?? string.Empty;
                    doc.Info.Title = "C2 CPS DATA";
                    for (int i = 0; i < dataCPS.Count; i++)
                    {
                        lednumber = dataCPS[i].LedNumber ?? string.Empty;
                        workno = dataCPS[0].WorkNo ?? "";
                        documentno = string.Format("ที่ RC_บค {1}_{0}", string.IsNullOrEmpty(workno) ? lednumber : workno, setdata.FestNo);
                        PdfPage page = doc.AddPage();
                        page.Size = PdfSharp.PageSize.A4;
                        page.Orientation = PdfSharp.PageOrientation.Portrait;

                        XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsUnit.Centimeter);
                        XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                        XFont font = new XFont("thsarabun", 16, XFontStyleEx.Bold, options);
                        XFont font_h_b = new XFont("thsarabun", 0.28, XFontStyleEx.Bold, options);
                        XFont font_h = new XFont("thsarabun", 0.28, XFontStyleEx.Regular, options);
                        XFont font_normalbold = new XFont("thsarabun", 0.493, XFontStyleEx.Bold, options);
                        XFont font_normal = new XFont("thsarabun", 0.493, XFontStyleEx.Regular, options);
                        XFont font_normal_small = new XFont("thsarabun", 0.35, XFontStyleEx.Regular, options);
                        XFont font_normal_install_bold = new XFont("thsarabun", 0.39, XFontStyleEx.Bold, options);
                        XFont font_normal_small_under = new XFont("thsarabun", 0.42, XFontStyleEx.Underline, options);
                        XFont font_bold_small_under = new XFont("thsarabun", 0.42, XFontStyleEx.Bold | XFontStyleEx.Underline, options);
                        XFont font_small_bold = new XFont("thsarabun", 0.35, XFontStyleEx.Bold, options);

                        #region Header
                        XImage image = XImage.FromFile(string.Format("{0}{1}", Application.StartupPath, "Images\\ktc_logo.png"));
                        gfx.DrawImage(image, 2.65, 1.25, 2.47, 1.8);
                        gfx.DrawString("บริษัท บัตรกรุงไทย จำกัด (มหาชน)", font_h_b, XBrushes.Gray, new XRect(6.17, 1.23, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("591 อาคารสมัชชาวาณิช 2 ชั้น 14 ถนนสุขุมวิท แขวงคลองตันเหนือ เขตวัฒนา กรุงเทพฯ 10110", font_h, XBrushes.Gray, new XRect(6.17, 1.52, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("โทร: 02 123 5100 โทรสาร: 02 123 5190", font_h, XBrushes.Gray, new XRect(6.17, 1.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("ทะเบียนเลขที่ 0107545000110", font_h, XBrushes.Gray, new XRect(6.17, 2.08, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                        gfx.DrawString("Krungthai Card Public Company Limited", font_h_b, XBrushes.Gray, new XRect(6.17, 2.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("591 United Business Center II, 14FL., Sukhumvit 33 Rd, North Klongton, Wattana, Bangkok 10110", font_h, XBrushes.Gray, new XRect(6.17, 2.75, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("Tel: 02 123 5100 Fax: 02 123 5190", font_h, XBrushes.Gray, new XRect(6.17, 3.03, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawString(documentno, font_normalbold, XBrushes.Black, new XRect(2.5, 3.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(string.Format("วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.FestDate ?? "")), font_normalbold, XBrushes.Black, new XRect(-2.5, 4.63, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        #endregion
                        #region Content 1
                        gfx.DrawString("เรื่อง", font_normalbold, XBrushes.Black, new XRect(2.5, 5.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("ผลการพิจารณาข้อตกลงการประนอมหนี้ตามข้อเสนอของผู้ร้อง", font_normal, XBrushes.Black, new XRect(4.0, 5.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("อ้างถึง", font_normalbold, XBrushes.Black, new XRect(2.5, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawString("บัญชีเคทีซี ของ", font_normal, XBrushes.Black, new XRect(4.0, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(string.Format("{0}", dataCPS[i].CustomerName), font_normalbold, XBrushes.Black, new XRect(6.0, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawString("หมายเลข", font_normal, XBrushes.Black, new XRect(4.0, 5.47 + spaceH + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(string.Format("{0}", dataCPS[i].CardNo), font_normalbold, XBrushes.Black, new XRect(5.25, 5.47 + spaceH + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                        string text1 = "ตามที่ บริษัท บัตรกรุงไทย จำกัด ( มหาชน ) หรือ “เคทีซี” ได้ยกเลิกสมาชิกพร้อมเรียกเก็บหนี้คืนทั้งหมด/";
                        XRect rect1 = new XRect(4.0, 8.11, 14.5, 10);
                        DrawJustifiedText(gfx, text1, font_normal, rect1);

                        string text2 = "ได้ดำเนินคดีกับผู้ร้อง ซึ่งศาลได้มีคำพิพากษาถึงที่สุดแล้ว และ ผู้ร้องยื่นคำร้องขอไกล่เกลี่ยข้อพิพาทในชั้นบังคับคดี ในคดีของ";
                        XRect rect2 = new XRect(2.5, 8.11 + spaceH, 16, 10);
                        DrawJustifiedText(gfx, text2, font_normal, rect2);

                        string courtname = string.Format("{0}", dataCPS[i].CourtName);
                        gfx.DrawString(courtname, font_normalbold, XBrushes.Black, new XRect(2.5, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        string textredno = "หมายเลขคดีแดงที่";
                        gfx.DrawString(textredno, font_normal, XBrushes.Black, new XRect(gfx.MeasureString(courtname, font_normalbold).Width + 2.8, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                        double widthdata = gfx.MeasureString(textredno, font_normal).Width + gfx.MeasureString(courtname, font_normalbold).Width;
                        string rednumber = string.Format("{0}", dataCPS[i].RedNo);
                        gfx.DrawString(rednumber, font_normalbold, XBrushes.Black, new XRect(widthdata + 3, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        string textktc = "ระหว่าง บริษัท บัตรกรุงไทย จำกัด ( มหาชน ) โจทก์ กับ";
                        gfx.DrawString(textktc, font_normal, XBrushes.Black, new XRect(-2.5, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);

                        gfx.DrawString(string.Format("{0} จำเลย", dataCPS[i].CustomerName), font_normal, XBrushes.Black, new XRect(2.5, 8.9 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        string text3 = "ต่อมาผู้ร้องได้ขอลดหย่อนภาระหนี้ เพื่อชำระหนี้ให้กับบริษัทฯ เป็นการเสร็จสิ้นตามความสามารถการชำระหนี้";
                        XRect rect3 = new XRect(4.0, 10.5, 14.5, 10);
                        DrawJustifiedText(gfx, text3, font_normal, rect3);

                        gfx.DrawString("ของผู้ร้องและบริษัทฯ ได้อนุมัติลดหย่อนภาระหนี้ โดยมีเงื่อนไขให้รับชำระหนี้เสร็จสิ้น ดังนี้", font_normal, XBrushes.Black, new XRect(2.5, 10.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        #endregion
                        #region TABLE

                        XBrush backgroundBrush = XBrushes.LightGray;
                        XBrush whiteBrush = XBrushes.Gainsboro;
                        //Table HD
                        XPen pen = new XPen(XColors.Black, 0.01);
                        gfx.DrawRectangle(pen, backgroundBrush, 2.2, 11.50, 4.28, 1.76);
                        gfx.DrawRectangle(pen, backgroundBrush, 6.48, 11.50, 2.7, 1.76);
                        gfx.DrawRectangle(pen, backgroundBrush, 9.18, 11.50, 2.9, 1.76);

                        gfx.DrawRectangle(pen, backgroundBrush, 12.08, 11.50, 2.20, 1.13);
                        gfx.DrawRectangle(pen, backgroundBrush, 14.28, 11.50, 2.30, 1.13);
                        gfx.DrawRectangle(pen, backgroundBrush, 16.48, 11.50, 2.30, 1.13);

                        gfx.DrawRectangle(pen, backgroundBrush, 12.08, 12.63, 2.20, 0.63);
                        gfx.DrawRectangle(pen, backgroundBrush, 14.28, 12.63, 2.30, 0.63);
                        gfx.DrawRectangle(pen, backgroundBrush, 16.48, 12.63, 2.30, 0.63);

                        ////Detail

                        gfx.DrawRectangle(pen, 2.2, 13.26, 4.28, 0.63);
                        gfx.DrawRectangle(pen, 2.2, 13.89, 4.28, 0.63);
                        gfx.DrawRectangle(pen, 2.2, 14.52, 4.28, 0.63);
                        gfx.DrawRectangle(pen, 2.2, 15.15, 4.28, 0.63);
                        gfx.DrawRectangle(pen, 2.2, 15.78, 4.28, 0.63);

                        gfx.DrawRectangle(pen, 6.48, 13.26, 2.7, 0.63);
                        gfx.DrawRectangle(pen, 6.48, 13.89, 2.7, 0.63);
                        gfx.DrawRectangle(pen, 6.48, 14.52, 2.7, 0.63);
                        gfx.DrawRectangle(pen, 6.48, 15.15, 2.7, 0.63);
                        gfx.DrawRectangle(pen, 6.48, 15.78, 2.7, 0.63);

                        gfx.DrawRectangle(pen, whiteBrush, 9.18, 13.26, 2.9, 0.63);
                        gfx.DrawRectangle(pen, 9.18, 13.89, 2.9, 0.63);
                        gfx.DrawRectangle(pen, 9.18, 14.52, 2.9, 0.63);
                        gfx.DrawRectangle(pen, 9.18, 15.15, 2.9, 0.63);
                        gfx.DrawRectangle(pen, 9.18, 15.78, 2.9, 0.63);


                        gfx.DrawRectangle(pen, whiteBrush, 12.08, 13.26, 2.20, 0.63);
                        gfx.DrawRectangle(pen, 12.08, 13.89, 2.20, 0.63);
                        gfx.DrawRectangle(pen, 12.08, 14.52, 2.20, 0.63);
                        gfx.DrawRectangle(pen, 12.08, 15.15, 2.20, 0.63);
                        gfx.DrawRectangle(pen, 12.08, 15.78, 2.20, 0.63);

                        gfx.DrawRectangle(pen, whiteBrush, 14.28, 13.26, 2.20, 0.63);
                        gfx.DrawRectangle(pen, 14.28, 13.89, 2.20, 0.63);
                        gfx.DrawRectangle(pen, 14.28, 14.52, 2.20, 0.63);
                        gfx.DrawRectangle(pen, 14.28, 15.15, 2.20, 0.63);
                        gfx.DrawRectangle(pen, 14.28, 15.78, 2.20, 0.63);

                        gfx.DrawRectangle(pen, whiteBrush, 16.48, 13.26, 2.30, 0.63);
                        gfx.DrawRectangle(pen, 16.48, 13.89, 2.30, 0.63);
                        gfx.DrawRectangle(pen, 16.48, 14.52, 2.30, 0.63);
                        gfx.DrawRectangle(pen, 16.48, 15.15, 2.30, 0.63);
                        gfx.DrawRectangle(pen, 16.48, 15.78, 2.30, 0.63);

                        gfx.DrawString("เงื่อนไขการชำระ", font_normal_install_bold, XBrushes.Black, new XRect(3.5, 12.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("จำนวนเงินที่ชำระ", font_normal_install_bold, XBrushes.Black, new XRect(6.85, 12.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("จำนวนเงินที่ชำระ", font_normal_install_bold, XBrushes.Black, new XRect(9.7, 12.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("ต่องวด", font_normal_install_bold, XBrushes.Black, new XRect(10.3, 12.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(12.85, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(12.4, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(12.45, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(15.05, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(14.6, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(14.65, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(17.25, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(16.8, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(16.85, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        double diameter = 0.35; // เส้นผ่านศูนย์กลางของวงกลมในเซนติเมตร
                        gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 13.56 - diameter / 2, diameter, diameter);
                        gfx.DrawString("ชำระปิดบัญชีคราวเดียว", font_normal_small, XBrushes.Black, new XRect(2.85, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(10.6, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(13.1, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(15.3, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(17.6, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 14.19 - diameter / 2, diameter, diameter);
                        gfx.DrawString("ผ่อนชำระไม่เกิน          6          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 13.98, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                        gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 14.82 - diameter / 2, diameter, diameter);
                        gfx.DrawString("ผ่อนชำระไม่เกิน        12          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 14.61, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                        gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 15.45 - diameter / 2, diameter, diameter);
                        gfx.DrawString("ผ่อนชำระไม่เกิน        24          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 15.24, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                        gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 16.08 - diameter / 2, diameter, diameter);
                        gfx.DrawString("ผ่อนชำระไม่เกิน ......................... งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 15.87, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);



                        #endregion
                        #region Content 2
                        gfx.DrawString("โดยเริ่มชำระงวดแรกภายในวันที่ .................................. งวดต่อไปชำระ ทุกๆ วันที่ ...................... ของทุกเดือน โดย", font_normal, XBrushes.Black, new XRect(4.0, 16.65, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("ชำระให้เสร็จสิ้นภายในวันที่ ...............................................ทั้งนี้บริษัทฯ อาจจะมีการแจ้งเตือนก่อนวันถึงกำหนดชำระผ่าน SMS", font_normal, XBrushes.Black, new XRect(2.5, 16.65 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        string custtel = string.Format("{0}", dataCPS[i].CustomerTel);
                        string textcusttel = "ตามหมายเลขโทรศัพท์ที่ท่านได้ให้ไว้แก่บริษัทฯ ครั้งหลังสุด";
                        string textcusttelother = "หรือ ...................................................";
                        gfx.DrawString(textcusttel, font_normal, XBrushes.Black, new XRect(2.5, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(custtel, font_normalbold, XBrushes.Black, new XRect(gfx.MeasureString(textcusttel, font_normal).Width + 2.65, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        double withcustel = gfx.MeasureString(textcusttel, font_normal).Width + gfx.MeasureString(custtel, font_normalbold).Width;
                        gfx.DrawString(textcusttelother, font_normal, XBrushes.Black, new XRect(withcustel + 2.8, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        string text4 = "ภายในกำหนดข้างต้นหากผู้ร้องชำระหนี้ดังกล่าวให้กับบริษัทฯเป็นที่ครบถ้วนเรียบร้อยแล้ว บริษัทฯจะถือการชำระ";
                        XRect rect4 = new XRect(4.0, 18.90, 14.5, 10);
                        DrawJustifiedText(gfx, text4, font_normal, rect4);

                        string text5 = "หนี้ตามจำนวนดังกล่าวเป็นการชำระหนี้เสร็จสิ้นตามที่บริษัทฯอนุมัติ โดยบริษัทฯจักได้ดำเนินการปรับปรุงบัญชีของผู้ร้องต่อไป";
                        XRect rect5 = new XRect(2.5, 18.90 + spaceH, 16, 10);
                        DrawJustifiedText(gfx, text5, font_normal, rect5);

                        string txtcontact = "ติดต่อพนักงานผู้รับผิดชอบ ";
                        string txtcollecttor = string.Format("{0} ", dataCPS[i].CollectorName);
                        string txttel = "โทร. ";
                        string txtcolltel = string.Format("{0}", dataCPS[i].CollectorTel);
                        gfx.DrawString(txtcontact, font_normal_small_under, XBrushes.Black, new XRect(2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(txtcollecttor, font_bold_small_under, XBrushes.Black, new XRect(gfx.MeasureString(txtcontact, font_normal_small_under).Width + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        double withcontact = gfx.MeasureString(txtcontact, font_normal_small_under).Width + gfx.MeasureString(txtcollecttor, font_bold_small_under).Width;
                        gfx.DrawString(txttel, font_normal_small_under, XBrushes.Black, new XRect(withcontact + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(txtcolltel, font_bold_small_under, XBrushes.Black, new XRect(withcontact + gfx.MeasureString(txttel, font_normal_small_under).Width + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawString("หมายเหตุ", font_normal_small, XBrushes.Black, new XRect(2.5, 21.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("1. กรณีมีการชำระหนี้ข้างต้นถือเป็นการชำระหนี้เสร็จสิ้นเฉพาะบัตรเลขที่ดังกล่าวข้างต้นเท่านั้น ไม่รวมถึงภาระหนี้อื่นของ", font_normal_small, XBrushes.Black, new XRect(4.0, 21.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("บมจ.บัตรกรุงไทย (ถ้ามี)", font_normal_small, XBrushes.Black, new XRect(4.0, 21.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("2. หากผู้ร้องผิดนัด ไม่ชำระหนี้ตามข้อตกลงไม่ว่าข้อหนึ่งข้อใด ถือว่าข้อตกลงไหล่เกลี่ยเป็นอันยกเลิก และบมจ.บัตรกรุงไทยจะดำเนินการ", font_normal_small, XBrushes.Black, new XRect(4.0, 22.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("บังคับคดีตามกฎหมายต่อไป โดยกลับไปคิดยอดหนี้ตามคำพิพากษาของศาลส่วนเงินที่ผู้ร้องชำระเข้ามาตามข้อตกลงไกล่เกลี่ย ให้เป็นส่วน", font_normal_small, XBrushes.Black, new XRect(4.0, 22.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("หนึ่งของการชำระหนี้ตามคำพิพากษา (ถ้ามี)", font_normal_small, XBrushes.Black, new XRect(4.0, 23.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                        gfx.DrawString("ลงชื่อ .......................................................... ผู้ร้องขอไกล่เกลี่ย", font_normal, XBrushes.Black, new XRect(2.5, 25.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("ลงชื่อ .......................................................... ผู้แทนโจทก์", font_normal, XBrushes.Black, new XRect(-2.5, 25.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        gfx.DrawString("(                                              )", font_normal, XBrushes.Black, new XRect(3.1, 25.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString("(                                              )", font_normal, XBrushes.Black, new XRect(-3.85, 25.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        gfx.DrawString(string.Format("ลำดับกรม {0}", dataCPS[i].LedNumber), font_small_bold, XBrushes.Black, new XRect(-2.5, 27.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        #endregion

                    }
                }               
            }
            string file_name = string.Empty;
            if (typerange == "L") file_name = string.Format("C2_LEDNUMBER_{0}To{1}.pdf",f_lednumber,l_lednumber);
            if (typerange == "W") file_name = string.Format("C2_WORKNO_{0}To{1}.pdf", f_workno, l_workno);
            string fullfilename = Path.Combine(C2PathFile, file_name);
            if (Path.Exists(C2PathFile))
            {
                doc.Save(fullfilename);
                return true;
            }
            return false;
        }
        public bool doCreateCPSTableReportMerge(ArrayList cpscardlist, SettingData setdata,string typerange)
        {
            GlobalFontSettings.FontResolver = new FileFontResolver();
            PdfDocument doc = new PdfDocument();
            List<DataCPSCard> dataCPS = new List<DataCPSCard>();
            double sumjudgmentAmnt = 0;
            double sumcapitalAmnt = 0;
            double sumdeptAmnt = 0;
            double sumaccCloseAmnt = 0;
            double sumaccClose6Amnt = 0;
            double suminstallment6Amnt = 0;
            double sumaccClose12Amnt = 0;
            double suminstallment12Amnt = 0;
            double sumaccClose24Amnt = 0;
            double suminstallment24Amnt = 0;
            string f_lednumber = string.Empty;
            string l_lednumber = string.Empty;
            string f_workno = string.Empty;
            string l_workno = string.Empty;
            for (int n = 0; n < cpscardlist.Count; n++)
            {
                var cpscard = cpscardlist[n];
                if (cpscard != null)
                {
                    dataCPS = (List<DataCPSCard>)cpscard;
                    if (n == 0) f_lednumber = dataCPS[0].LedNumber ?? string.Empty;
                    if (n == cpscardlist.Count - 1) l_lednumber = dataCPS[0].LedNumber ?? string.Empty;
                    if (n == 0) f_workno = dataCPS[0].WorkNo ?? string.Empty;
                    if (n == cpscardlist.Count - 1) l_workno = dataCPS[0].WorkNo ?? string.Empty;

                    doc.Info.Title = "CPS Table";
                    PdfPage page = doc.AddPage();
                    page.Size = PdfSharp.PageSize.A4;
                    page.Orientation = PdfSharp.PageOrientation.Landscape;

                    XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsUnit.Centimeter);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                    XFont font_title_bold = new XFont("thsarabun", 0.56, XFontStyleEx.Bold, options);
                    XFont font_normalbold = new XFont("thsarabun", 0.49, XFontStyleEx.Bold, options);
                    XFont font_normal = new XFont("thsarabun", 0.49, XFontStyleEx.Regular, options);
                    XFont font_normal_underline = new XFont("thsarabun", 0.49, XFontStyleEx.Underline, options);
                    XFont font_normal_small = new XFont("thsarabun", 0.38, XFontStyleEx.Regular, options);
                    XFont font_bold_small = new XFont("thsarabun", 0.38, XFontStyleEx.Bold, options);
                    XPen pen = new XPen(XColors.Black, 0.03);

                    XImage image = XImage.FromFile(string.Format("{0}{1}", Application.StartupPath, @"Images/ktc_logo.png"));
                    gfx.DrawImage(image, 1.1, 1.72, 3.24, 2.33);

                    if (!string.IsNullOrEmpty(dataCPS[0].LedNumber))
                    {
                        BitMatrix Qrcode = qrcodeService.Encode(dataCPS[0].LedNumber ?? "", 1);
                        qrcodeService.DrawQrCode(gfx, new XPoint(25.6, 1.3), Qrcode, 0.10);
                    }
                    gfx.DrawString(string.Format("มหกรรมไกล่เกลี่ยชั้นบังคับคดี ครั้งที่ {0}", setdata.FestNo), font_title_bold, XBrushes.Black, new XRect(0, 1.72, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
                    gfx.DrawString(string.Format("{0}", setdata.FestName), font_title_bold, XBrushes.Black, new XRect(0, 2.4, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
                    gfx.DrawString(string.Format("วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.FestDate ?? "")), font_title_bold, XBrushes.Black, new XRect(0, 3.1, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

                    gfx.DrawString("ลำดับที่", font_normalbold, XBrushes.Black, new XRect(1.2, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].WorkNo) ? "-" : dataCPS[0].WorkNo), font_normal, XBrushes.Black, new XRect(2.3, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("ลำดับกรม", font_normalbold, XBrushes.Black, new XRect(-1, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
                    gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].LedNumber) ? "-" : dataCPS[0].LedNumber), font_normal, XBrushes.Black, new XRect(0.2, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

                    gfx.DrawString("เลขที่คดีดำ", font_normalbold, XBrushes.Black, new XRect(21.5, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].BlackNo) ? "-" : dataCPS[0].BlackNo), font_normal, XBrushes.Black, new XRect(23.1, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("ชื่อ", font_normalbold, XBrushes.Black, new XRect(1.2, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", dataCPS[0].CustomerName), font_normal, XBrushes.Black, new XRect(1.7, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString(string.Format("ภาระหนี้ ณ วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.DateAtCalulate ?? "")), font_normalbold, XBrushes.Black, new XRect(-0.3, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

                    gfx.DrawString("เลขที่คดีแดง", font_normalbold, XBrushes.Black, new XRect(21.5, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", dataCPS[0].RedNo), font_normal, XBrushes.Black, new XRect(23.25, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("วันพิพากษา", font_normalbold, XBrushes.Black, new XRect(21.5, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", datehelper.doGetShortDateTHFromDBToPDF(dataCPS[0].JudgeDate ?? "")), font_normal, XBrushes.Black, new XRect(23.2, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    double diameter = 0.48; // เส้นผ่านศูนย์กลางของวงกลมในเซนติเมตร
                    gfx.DrawEllipse(pen, XBrushes.White, 1.6 - diameter / 2, 6.75 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ตกลงรับเงื่อนไข", font_normal, XBrushes.Black, new XRect(2.1, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 5.6 - diameter / 2, 6.75 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ไม่ตกลงรับเงื่อนไข  เนื่องจาก ............................................................", font_normal, XBrushes.Black, new XRect(6.1, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    #region TABLE
                    #region HEADER
                    XPen pentable = new XPen(XColors.Black, 0.015);
                    gfx.DrawRectangle(pentable, 1.29, 7.4, 0.85, 2.1);
                    gfx.DrawRectangle(pentable, 2.14, 7.4, 3.58, 2.1);
                    gfx.DrawRectangle(pentable, 5.72, 7.4, 2.61, 2.1);
                    gfx.DrawRectangle(pentable, 8.33, 7.4, 2.03, 2.1);
                    gfx.DrawRectangle(pentable, 10.36, 7.4, 2.47, 2.1);
                    gfx.DrawRectangle(pentable, 12.83, 7.4, 15.72, 2.1);
                    gfx.DrawString("ที่", font_normal, XBrushes.Black, new XRect(1.62, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("หมายเลขบัตร", font_normal, XBrushes.Black, new XRect(3.0, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("พิพากษาทั้งสิ้น", font_normal, XBrushes.Black, new XRect(6.1, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("เงินต้น", font_normal, XBrushes.Black, new XRect(8.9, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ภาระหนี้", font_normal, XBrushes.Black, new XRect(11.0, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ปัจจุบัน", font_normal, XBrushes.Black, new XRect(11.1, 8.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("เงื่อนไขในการชำระ", font_normal, XBrushes.Black, new XRect(19.4, 7.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                    gfx.DrawRectangle(pentable, 12.83, 8.1, 2.7, 0.7);
                    gfx.DrawRectangle(pentable, 12.83, 8.8, 2.7, 0.7);
                    gfx.DrawEllipse(pen, XBrushes.White, 13.3 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ปิดงวดเดียว", font_normal, XBrushes.Black, new XRect(13.75, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(13.65, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawRectangle(pentable, 15.53, 8.1, 4.34, 0.7);
                    gfx.DrawRectangle(pentable, 15.53, 8.8, 2.17, 0.7);
                    gfx.DrawRectangle(pentable, 17.70, 8.8, 2.17, 0.7);
                    gfx.DrawEllipse(pen, XBrushes.White, 16.0 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อน 6 งวด", font_normal, XBrushes.Black, new XRect(17.2, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(16.10, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(18.4, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawRectangle(pentable, 19.87, 8.1, 4.34, 0.7);
                    gfx.DrawRectangle(pentable, 19.87, 8.8, 2.17, 0.7);
                    gfx.DrawRectangle(pentable, 22.04, 8.8, 2.17, 0.7);
                    gfx.DrawEllipse(pen, XBrushes.White, 20.4 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อน 12 งวด", font_normal, XBrushes.Black, new XRect(21.5, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(20.5, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(22.8, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawRectangle(pentable, 24.21, 8.1, 4.34, 0.7);
                    gfx.DrawRectangle(pentable, 24.21, 8.8, 2.17, 0.7);
                    gfx.DrawRectangle(pentable, 26.38, 8.8, 2.17, 0.7);
                    gfx.DrawEllipse(pen, XBrushes.White, 24.8 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อน 24 งวด", font_normal, XBrushes.Black, new XRect(25.9, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(24.9, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(27.22, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    #endregion
                    #region Detail
                    double ypos_strat = 9.5;
                    double ypostext_strat = 9.55;
                    string dat = "-";

                    for (int i = 0; i < 7; i++)
                    {
                        if (i != 6)
                        {
                            gfx.DrawRectangle(pentable, 1.29, ypos_strat, 0.85, 0.7);
                            gfx.DrawString(string.Format("{0}", i + 1), font_normal, XBrushes.Black, new XRect(1.62, ypos_strat + 0.05, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        gfx.DrawRectangle(pentable, 2.14, ypos_strat, 3.58, 0.7);
                        gfx.DrawRectangle(pentable, 5.72, ypos_strat, 2.61, 0.7);
                        gfx.DrawRectangle(pentable, 8.33, ypos_strat, 2.03, 0.7);
                        gfx.DrawRectangle(pentable, 10.36, ypos_strat, 2.47, 0.7);
                        gfx.DrawRectangle(pentable, 12.83, ypos_strat, 2.7, 0.7);
                        gfx.DrawRectangle(pentable, 15.53, ypos_strat, 2.17, 0.7);
                        gfx.DrawRectangle(pentable, 17.70, ypos_strat, 2.17, 0.7);
                        gfx.DrawRectangle(pentable, 19.87, ypos_strat, 2.17, 0.7);
                        gfx.DrawRectangle(pentable, 22.04, ypos_strat, 2.17, 0.7);
                        gfx.DrawRectangle(pentable, 24.21, ypos_strat, 2.17, 0.7);
                        gfx.DrawRectangle(pentable, 26.38, ypos_strat, 2.17, 0.7);

                        if (i < dataCPS.Count) // SET VALE FROM DB
                        {
                            #region calculate value
                            double judgmentAmnt = dataCPS[i].JudgmentAmnt;
                            double capitalAmnt = dataCPS[i].PrincipleAmnt; //CapitalAmnt;
                            double deptAmnt = dataCPS[i].DeptAmnt;
                            double accCloseAmnt = dataCPS[i].AccCloseAmnt;
                            double accClose6Amnt = dataCPS[i].AccClose6Amnt;
                            double installment6Amnt = dataCPS[i].Installment6Amnt;
                            double accClose12Amnt = dataCPS[i].AccClose12Amnt;
                            double installment12Amnt = dataCPS[i].Installment12Amnt;
                            double accClose24Amnt = dataCPS[i].AccClose24Amnt;
                            double installment24Amnt = dataCPS[i].Installment24Amnt;
                            string cardno = string.Format("{0}", dataCPS[i].CardNo);

                            if (dataCPS[i].Maxmonth < 6)
                            {
                                accClose6Amnt = 0;
                                installment6Amnt = 0;
                            }
                            if (dataCPS[i].Maxmonth < 12)
                            {
                                accClose12Amnt = 0;
                                installment12Amnt = 0;
                            }
                            if (dataCPS[i].Maxmonth < 24)
                            {
                                accClose24Amnt = 0;
                                installment24Amnt = 0;
                            }

                            sumjudgmentAmnt = sumjudgmentAmnt + judgmentAmnt;
                            sumcapitalAmnt = sumcapitalAmnt + capitalAmnt;
                            sumdeptAmnt = sumdeptAmnt + deptAmnt;
                            sumaccCloseAmnt = sumaccCloseAmnt + accCloseAmnt;
                            sumaccClose6Amnt = sumaccClose6Amnt + accClose6Amnt;
                            suminstallment6Amnt = suminstallment6Amnt + installment6Amnt;
                            sumaccClose12Amnt = sumaccClose12Amnt + accClose12Amnt;
                            suminstallment12Amnt = suminstallment12Amnt + installment12Amnt;
                            sumaccClose24Amnt = sumaccClose24Amnt + accClose24Amnt;
                            suminstallment24Amnt = suminstallment24Amnt + installment24Amnt;
                            #endregion

                            #region SET Value
                            if (!string.IsNullOrEmpty(cardno))
                            {
                                gfx.DrawString(cardno, font_normal, XBrushes.Black, new XRect(2.4, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(3.9, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (judgmentAmnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", judgmentAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-21.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (capitalAmnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", capitalAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-19.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (deptAmnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", deptAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-17.0, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (accCloseAmnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", accCloseAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-14.3, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (accClose6Amnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", accClose6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-12.1, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (installment6Amnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", installment6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-9.95, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (accClose12Amnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", accClose12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-7.80, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (installment12Amnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", installment12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-5.58, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (accClose24Amnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", accClose24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-3.4, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }

                            if (installment24Amnt > 0)
                            {
                                gfx.DrawString(string.Format("{0}", installment24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-1.25, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }
                            #endregion

                        }
                        else
                        {
                            #region SET SUM VALUE
                            if (i == 6)
                            {
                                gfx.DrawString("รวม", font_normal_underline, XBrushes.Black, new XRect(3.69, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                if (sumjudgmentAmnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", sumjudgmentAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-21.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }

                                if (sumcapitalAmnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", sumcapitalAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-19.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }

                                if (sumdeptAmnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", sumdeptAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-17.0, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }

                                if (sumaccCloseAmnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", sumaccCloseAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-14.3, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }

                                if (sumaccClose6Amnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", sumaccClose6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-12.1, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }

                                if (suminstallment6Amnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", suminstallment6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-9.95, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }

                                if (sumaccClose12Amnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", sumaccClose12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-7.80, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }
                                if (suminstallment12Amnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", suminstallment12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-5.58, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }
                                if (sumaccClose24Amnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", sumaccClose24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-3.4, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }
                                if (suminstallment24Amnt > 0)
                                {
                                    gfx.DrawString(string.Format("{0}", suminstallment24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-1.25, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                                }
                                else
                                {
                                    gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                }


                            }
                            else
                            {
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(3.9, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                                gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                            }
                            #endregion
                        }
                        ypostext_strat = ypostext_strat + 0.7;
                        ypos_strat = ypos_strat + 0.7;
                    }
                    #region TP LP
                    double y_point = 15.2;
                    List<string> datashow = CreateTPDP(dataCPS);
                    for (int m = 0; m < 6; m++)
                    {
                        gfx.DrawString(datashow[m], font_normal_small, XBrushes.Black, new XRect(12, y_point, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        y_point = y_point + 0.7;
                    }
                    #endregion
                    #endregion
                    #endregion
                    gfx.DrawString("สถานะทางคดี :", font_bold_small, XBrushes.Black, new XRect(2.14, 15.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", dataCPS[0].LegalStatus), font_normal_small, XBrushes.Black, new XRect(3.8, 15.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("หมายเหตุบังคับคดี :", font_bold_small, XBrushes.Black, new XRect(2.14, 15.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].LegalExecRemark) ? "-" : dataCPS[0].LegalExecRemark), font_normal_small, XBrushes.Black, new XRect(4.3, 15.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("วันที่ยึดทรัพย์/อายัดเงินเดือน :", font_bold_small, XBrushes.Black, new XRect(2.14, 16.6, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", datehelper.doGetDateThaiFromDBToPDF(dataCPS[0].LegalExecDate ?? "-")), font_normal_small, XBrushes.Black, new XRect(5.4, 16.6, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("ตรวจสอบ : _________________________________", font_normalbold, XBrushes.Black, new XRect(2.14, 18.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("ผู้เจรจา : __________________________________", font_normalbold, XBrushes.Black, new XRect(18.4, 16.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString(string.Format("เจ้าหน้าที่ : | {0} | {1} | เบอร์ติดต่อ : {2}", dataCPS[0].CollectorName, string.IsNullOrEmpty(dataCPS[0].CollectorTeam) ? "-" : dataCPS[0].CollectorTeam, dataCPS[0].CollectorTel), font_normalbold, XBrushes.Black, new XRect(-3.5, 18.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);

                }
            }
            string file_name = string.Empty;
            if (typerange == "L") file_name = string.Format("Table_LEDNUMBER_{0}To{1}.pdf", f_lednumber, l_lednumber);
            if (typerange == "W") file_name = string.Format("Table_WORKNO_{0}To{1}.pdf", f_workno, l_workno);
            string fullfilename = Path.Combine(TablePathFile, file_name);
            if (Path.Exists(TablePathFile))
            {
                doc.Save(fullfilename);
                return true;
            }           
            return false;
        }

        public bool doUserCreateC2PDFReport(List<DataCPSCard> dataCPS, SettingData setdata,string queueno)
        {
            GlobalFontSettings.FontResolver = new FileFontResolver();
            PdfDocument doc = new PdfDocument();
            double spaceH = 0.6;
            string lednumber = string.Empty;

            doc.Info.Title = "C2 CPS DATA";
            for (int i = 0; i < dataCPS.Count; i++)
            {
                for (int dup = 0; dup < 2; dup++)
                {
                    lednumber = dataCPS[i].LedNumber ?? string.Empty;
                    workno = dataCPS[0].WorkNo ?? "";
                    documentno = string.Format("ที่ RC_บค {1}_{0}", string.IsNullOrEmpty(workno) ? lednumber : workno, setdata.FestNo);
                    PdfPage page = doc.AddPage();
                    page.Size = PdfSharp.PageSize.A4;
                    page.Orientation = PdfSharp.PageOrientation.Portrait;

                    XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsUnit.Centimeter);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                    XFont font = new XFont("thsarabun", 16, XFontStyleEx.Bold, options);
                    XFont font_h_b = new XFont("thsarabun", 0.28, XFontStyleEx.Bold, options);
                    XFont font_h = new XFont("thsarabun", 0.28, XFontStyleEx.Regular, options);
                    XFont font_normalbold = new XFont("thsarabun", 0.493, XFontStyleEx.Bold, options);
                    XFont font_normal = new XFont("thsarabun", 0.493, XFontStyleEx.Regular, options);
                    XFont font_normal_small = new XFont("thsarabun", 0.35, XFontStyleEx.Regular, options);
                    XFont font_normal_install_bold = new XFont("thsarabun", 0.39, XFontStyleEx.Bold, options);
                    XFont font_normal_small_under = new XFont("thsarabun", 0.42, XFontStyleEx.Underline, options);
                    XFont font_bold_small_under = new XFont("thsarabun", 0.42, XFontStyleEx.Bold | XFontStyleEx.Underline, options);
                    XFont font_small_bold = new XFont("thsarabun", 0.35, XFontStyleEx.Bold, options);

                    #region Header
                    XImage image = XImage.FromFile(string.Format("{0}{1}", Application.StartupPath, "Images\\ktc_logo.png"));
                    gfx.DrawImage(image, 2.65, 1.25, 2.47, 1.8);
                    gfx.DrawString("บริษัท บัตรกรุงไทย จำกัด (มหาชน)", font_h_b, XBrushes.Gray, new XRect(6.17, 1.23, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("591 อาคารสมัชชาวาณิช 2 ชั้น 14 ถนนสุขุมวิท แขวงคลองตันเหนือ เขตวัฒนา กรุงเทพฯ 10110", font_h, XBrushes.Gray, new XRect(6.17, 1.52, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("โทร: 02 123 5100 โทรสาร: 02 123 5190", font_h, XBrushes.Gray, new XRect(6.17, 1.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ทะเบียนเลขที่ 0107545000110", font_h, XBrushes.Gray, new XRect(6.17, 2.08, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    gfx.DrawString("Krungthai Card Public Company Limited", font_h_b, XBrushes.Gray, new XRect(6.17, 2.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("591 United Business Center II, 14FL., Sukhumvit 33 Rd, North Klongton, Wattana, Bangkok 10110", font_h, XBrushes.Gray, new XRect(6.17, 2.75, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("Tel: 02 123 5100 Fax: 02 123 5190", font_h, XBrushes.Gray, new XRect(6.17, 3.03, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString(documentno, font_normalbold, XBrushes.Black, new XRect(2.5, 3.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.FestDate ?? "")), font_normalbold, XBrushes.Black, new XRect(-2.5, 4.63, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    #endregion
                    #region Content 1
                    gfx.DrawString("เรื่อง", font_normalbold, XBrushes.Black, new XRect(2.5, 5.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ผลการพิจารณาข้อตกลงการประนอมหนี้ตามข้อเสนอของผู้ร้อง", font_normal, XBrushes.Black, new XRect(4.0, 5.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("อ้างถึง", font_normalbold, XBrushes.Black, new XRect(2.5, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("บัญชีเคทีซี ของ", font_normal, XBrushes.Black, new XRect(4.0, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", dataCPS[i].CustomerName), font_normalbold, XBrushes.Black, new XRect(6.0, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("หมายเลข", font_normal, XBrushes.Black, new XRect(4.0, 5.47 + spaceH + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", dataCPS[i].CardNo), font_normalbold, XBrushes.Black, new XRect(5.25, 5.47 + spaceH + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                    string text1 = "ตามที่ บริษัท บัตรกรุงไทย จำกัด ( มหาชน ) หรือ “เคทีซี” ได้ยกเลิกสมาชิกพร้อมเรียกเก็บหนี้คืนทั้งหมด/";
                    XRect rect1 = new XRect(4.0, 8.11, 14.5, 10);
                    DrawJustifiedText(gfx, text1, font_normal, rect1);

                    string text2 = "ได้ดำเนินคดีกับผู้ร้อง ซึ่งศาลได้มีคำพิพากษาถึงที่สุดแล้ว และ ผู้ร้องยื่นคำร้องขอไกล่เกลี่ยข้อพิพาทในชั้นบังคับคดี ในคดีของ";
                    XRect rect2 = new XRect(2.5, 8.11 + spaceH, 16, 10);
                    DrawJustifiedText(gfx, text2, font_normal, rect2);

                    string courtname = string.Format("{0}", dataCPS[i].CourtName);
                    gfx.DrawString(courtname, font_normalbold, XBrushes.Black, new XRect(2.5, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string textredno = "หมายเลขคดีแดงที่";
                    gfx.DrawString(textredno, font_normal, XBrushes.Black, new XRect(gfx.MeasureString(courtname, font_normalbold).Width + 2.8, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                    double widthdata = gfx.MeasureString(textredno, font_normal).Width + gfx.MeasureString(courtname, font_normalbold).Width;
                    string rednumber = string.Format("{0}", dataCPS[i].RedNo);
                    gfx.DrawString(rednumber, font_normalbold, XBrushes.Black, new XRect(widthdata + 3, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string textktc = "ระหว่าง บริษัท บัตรกรุงไทย จำกัด ( มหาชน ) โจทก์ กับ";
                    gfx.DrawString(textktc, font_normal, XBrushes.Black, new XRect(-2.5, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);

                    gfx.DrawString(string.Format("{0} จำเลย", dataCPS[i].CustomerName), font_normal, XBrushes.Black, new XRect(2.5, 8.9 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string text3 = "ต่อมาผู้ร้องได้ขอลดหย่อนภาระหนี้ เพื่อชำระหนี้ให้กับบริษัทฯ เป็นการเสร็จสิ้นตามความสามารถการชำระหนี้";
                    XRect rect3 = new XRect(4.0, 10.5, 14.5, 10);
                    DrawJustifiedText(gfx, text3, font_normal, rect3);

                    gfx.DrawString("ของผู้ร้องและบริษัทฯ ได้อนุมัติลดหย่อนภาระหนี้ โดยมีเงื่อนไขให้รับชำระหนี้เสร็จสิ้น ดังนี้", font_normal, XBrushes.Black, new XRect(2.5, 10.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    #endregion
                    #region TABLE

                    XBrush backgroundBrush = XBrushes.LightGray;
                    XBrush whiteBrush = XBrushes.Gainsboro;
                    //Table HD
                    XPen pen = new XPen(XColors.Black, 0.01);
                    gfx.DrawRectangle(pen, backgroundBrush, 2.2, 11.50, 4.28, 1.76);
                    gfx.DrawRectangle(pen, backgroundBrush, 6.48, 11.50, 2.7, 1.76);
                    gfx.DrawRectangle(pen, backgroundBrush, 9.18, 11.50, 2.9, 1.76);

                    gfx.DrawRectangle(pen, backgroundBrush, 12.08, 11.50, 2.20, 1.13);
                    gfx.DrawRectangle(pen, backgroundBrush, 14.28, 11.50, 2.30, 1.13);
                    gfx.DrawRectangle(pen, backgroundBrush, 16.48, 11.50, 2.30, 1.13);

                    gfx.DrawRectangle(pen, backgroundBrush, 12.08, 12.63, 2.20, 0.63);
                    gfx.DrawRectangle(pen, backgroundBrush, 14.28, 12.63, 2.30, 0.63);
                    gfx.DrawRectangle(pen, backgroundBrush, 16.48, 12.63, 2.30, 0.63);

                    ////Detail

                    gfx.DrawRectangle(pen, 2.2, 13.26, 4.28, 0.63);
                    gfx.DrawRectangle(pen, 2.2, 13.89, 4.28, 0.63);
                    gfx.DrawRectangle(pen, 2.2, 14.52, 4.28, 0.63);
                    gfx.DrawRectangle(pen, 2.2, 15.15, 4.28, 0.63);
                    gfx.DrawRectangle(pen, 2.2, 15.78, 4.28, 0.63);

                    gfx.DrawRectangle(pen, 6.48, 13.26, 2.7, 0.63);
                    gfx.DrawRectangle(pen, 6.48, 13.89, 2.7, 0.63);
                    gfx.DrawRectangle(pen, 6.48, 14.52, 2.7, 0.63);
                    gfx.DrawRectangle(pen, 6.48, 15.15, 2.7, 0.63);
                    gfx.DrawRectangle(pen, 6.48, 15.78, 2.7, 0.63);

                    gfx.DrawRectangle(pen, whiteBrush, 9.18, 13.26, 2.9, 0.63);
                    gfx.DrawRectangle(pen, 9.18, 13.89, 2.9, 0.63);
                    gfx.DrawRectangle(pen, 9.18, 14.52, 2.9, 0.63);
                    gfx.DrawRectangle(pen, 9.18, 15.15, 2.9, 0.63);
                    gfx.DrawRectangle(pen, 9.18, 15.78, 2.9, 0.63);


                    gfx.DrawRectangle(pen, whiteBrush, 12.08, 13.26, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 12.08, 13.89, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 12.08, 14.52, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 12.08, 15.15, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 12.08, 15.78, 2.20, 0.63);

                    gfx.DrawRectangle(pen, whiteBrush, 14.28, 13.26, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 14.28, 13.89, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 14.28, 14.52, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 14.28, 15.15, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 14.28, 15.78, 2.20, 0.63);

                    gfx.DrawRectangle(pen, whiteBrush, 16.48, 13.26, 2.30, 0.63);
                    gfx.DrawRectangle(pen, 16.48, 13.89, 2.30, 0.63);
                    gfx.DrawRectangle(pen, 16.48, 14.52, 2.30, 0.63);
                    gfx.DrawRectangle(pen, 16.48, 15.15, 2.30, 0.63);
                    gfx.DrawRectangle(pen, 16.48, 15.78, 2.30, 0.63);

                    gfx.DrawString("เงื่อนไขการชำระ", font_normal_install_bold, XBrushes.Black, new XRect(3.5, 12.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("จำนวนเงินที่ชำระ", font_normal_install_bold, XBrushes.Black, new XRect(6.85, 12.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("จำนวนเงินที่ชำระ", font_normal_install_bold, XBrushes.Black, new XRect(9.7, 12.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ต่องวด", font_normal_install_bold, XBrushes.Black, new XRect(10.3, 12.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(12.85, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(12.4, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(12.45, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(15.05, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(14.6, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(14.65, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(17.25, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(16.8, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(16.85, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    double diameter = 0.35; // เส้นผ่านศูนย์กลางของวงกลมในเซนติเมตร
                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 13.56 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ชำระปิดบัญชีคราวเดียว", font_normal_small, XBrushes.Black, new XRect(2.85, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(10.6, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(13.1, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(15.3, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(17.6, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 14.19 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อนชำระไม่เกิน          6          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 13.98, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 14.82 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อนชำระไม่เกิน        12          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 14.61, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 15.45 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อนชำระไม่เกิน        24          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 15.24, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 16.08 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อนชำระไม่เกิน ......................... งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 15.87, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);



                    #endregion
                    #region Content 2
                    gfx.DrawString("โดยเริ่มชำระงวดแรกภายในวันที่ .................................. งวดต่อไปชำระ ทุกๆ วันที่ ...................... ของทุกเดือน โดย", font_normal, XBrushes.Black, new XRect(4.0, 16.65, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ชำระให้เสร็จสิ้นภายในวันที่ ...............................................ทั้งนี้บริษัทฯ อาจจะมีการแจ้งเตือนก่อนวันถึงกำหนดชำระผ่าน SMS", font_normal, XBrushes.Black, new XRect(2.5, 16.65 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string custtel = string.Format("{0}", dataCPS[i].CustomerTel);
                    string textcusttel = "ตามหมายเลขโทรศัพท์ที่ท่านได้ให้ไว้แก่บริษัทฯ ครั้งหลังสุด";
                    string textcusttelother = "หรือ ...................................................";
                    gfx.DrawString(textcusttel, font_normal, XBrushes.Black, new XRect(2.5, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(custtel, font_normalbold, XBrushes.Black, new XRect(gfx.MeasureString(textcusttel, font_normal).Width + 2.65, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    double withcustel = gfx.MeasureString(textcusttel, font_normal).Width + gfx.MeasureString(custtel, font_normalbold).Width;
                    gfx.DrawString(textcusttelother, font_normal, XBrushes.Black, new XRect(withcustel + 2.8, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string text4 = "ภายในกำหนดข้างต้นหากผู้ร้องชำระหนี้ดังกล่าวให้กับบริษัทฯเป็นที่ครบถ้วนเรียบร้อยแล้ว บริษัทฯจะถือการชำระ";
                    XRect rect4 = new XRect(4.0, 18.90, 14.5, 10);
                    DrawJustifiedText(gfx, text4, font_normal, rect4);

                    string text5 = "หนี้ตามจำนวนดังกล่าวเป็นการชำระหนี้เสร็จสิ้นตามที่บริษัทฯอนุมัติ โดยบริษัทฯจักได้ดำเนินการปรับปรุงบัญชีของผู้ร้องต่อไป";
                    XRect rect5 = new XRect(2.5, 18.90 + spaceH, 16, 10);
                    DrawJustifiedText(gfx, text5, font_normal, rect5);

                    string txtcontact = "ติดต่อพนักงานผู้รับผิดชอบ ";
                    string txtcollecttor = string.Format("{0} ", dataCPS[i].CollectorName);
                    string txttel = "โทร. ";
                    string txtcolltel = string.Format("{0}", dataCPS[i].CollectorTel);
                    gfx.DrawString(txtcontact, font_normal_small_under, XBrushes.Black, new XRect(2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(txtcollecttor, font_bold_small_under, XBrushes.Black, new XRect(gfx.MeasureString(txtcontact, font_normal_small_under).Width + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    double withcontact = gfx.MeasureString(txtcontact, font_normal_small_under).Width + gfx.MeasureString(txtcollecttor, font_bold_small_under).Width;
                    gfx.DrawString(txttel, font_normal_small_under, XBrushes.Black, new XRect(withcontact + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(txtcolltel, font_bold_small_under, XBrushes.Black, new XRect(withcontact + gfx.MeasureString(txttel, font_normal_small_under).Width + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("หมายเหตุ", font_normal_small, XBrushes.Black, new XRect(2.5, 21.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("1. กรณีมีการชำระหนี้ข้างต้นถือเป็นการชำระหนี้เสร็จสิ้นเฉพาะบัตรเลขที่ดังกล่าวข้างต้นเท่านั้น ไม่รวมถึงภาระหนี้อื่นของ", font_normal_small, XBrushes.Black, new XRect(4.0, 21.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("บมจ.บัตรกรุงไทย (ถ้ามี)", font_normal_small, XBrushes.Black, new XRect(4.0, 21.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("2. หากผู้ร้องผิดนัด ไม่ชำระหนี้ตามข้อตกลงไม่ว่าข้อหนึ่งข้อใด ถือว่าข้อตกลงไหล่เกลี่ยเป็นอันยกเลิก และบมจ.บัตรกรุงไทยจะดำเนินการ", font_normal_small, XBrushes.Black, new XRect(4.0, 22.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("บังคับคดีตามกฎหมายต่อไป โดยกลับไปคิดยอดหนี้ตามคำพิพากษาของศาลส่วนเงินที่ผู้ร้องชำระเข้ามาตามข้อตกลงไกล่เกลี่ย ให้เป็นส่วน", font_normal_small, XBrushes.Black, new XRect(4.0, 22.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("หนึ่งของการชำระหนี้ตามคำพิพากษา (ถ้ามี)", font_normal_small, XBrushes.Black, new XRect(4.0, 23.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("ลงชื่อ .......................................................... ผู้ร้องขอไกล่เกลี่ย", font_normal, XBrushes.Black, new XRect(2.5, 25.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ลงชื่อ .......................................................... ผู้แทนโจทก์", font_normal, XBrushes.Black, new XRect(-2.5, 25.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    gfx.DrawString("(                                              )", font_normal, XBrushes.Black, new XRect(3.1, 25.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("(                                              )", font_normal, XBrushes.Black, new XRect(-3.85, 25.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    gfx.DrawString(string.Format("ลำดับกรม {0}", dataCPS[i].LedNumber), font_small_bold, XBrushes.Black, new XRect(-2.5, 27.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    #endregion
                }
            }
            if (dataCPS.Count > 0)
            {
                string file_name = string.Format("C2_{0}_{1}.pdf",queueno,dataCPS[0].CustomerID);
                file_name = file_name.Replace("\n", "").Replace("\r", "").Replace("/", "").Replace(" ", "");
                string fullfilename = Path.Combine(C2PathFile, file_name);
                if (Path.Exists(C2PathFile))
                {
                    doc.Save(fullfilename);
                    return true;
                }
            }
            return false;
        }
        public bool doUserCreateCPSTableReport(List<DataCPSCard> dataCPS, SettingData setdata, string queueno)
        {
            GlobalFontSettings.FontResolver = new FileFontResolver();
            PdfDocument doc = new PdfDocument();

            double sumjudgmentAmnt = 0;
            double sumcapitalAmnt = 0;
            double sumdeptAmnt = 0;
            double sumaccCloseAmnt = 0;
            double sumaccClose6Amnt = 0;
            double suminstallment6Amnt = 0;
            double sumaccClose12Amnt = 0;
            double suminstallment12Amnt = 0;
            double sumaccClose24Amnt = 0;
            double suminstallment24Amnt = 0;

            doc.Info.Title = "CPS Table";
            PdfPage page = doc.AddPage();
            page.Size = PdfSharp.PageSize.A4;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsUnit.Centimeter);
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            XFont font_title_bold = new XFont("thsarabun", 0.56, XFontStyleEx.Bold, options);
            XFont font_normalbold = new XFont("thsarabun", 0.49, XFontStyleEx.Bold, options);
            XFont font_normal = new XFont("thsarabun", 0.49, XFontStyleEx.Regular, options);
            XFont font_normal_underline = new XFont("thsarabun", 0.49, XFontStyleEx.Underline, options);
            XFont font_normal_small = new XFont("thsarabun", 0.38, XFontStyleEx.Regular, options);
            XFont font_bold_small = new XFont("thsarabun", 0.38, XFontStyleEx.Bold, options);
            XPen pen = new XPen(XColors.Black, 0.03);

            XImage image = XImage.FromFile(string.Format("{0}{1}", Application.StartupPath, @"Images/ktc_logo.png"));
            gfx.DrawImage(image, 1.1, 1.72, 3.24, 2.33);

            if (!string.IsNullOrEmpty(dataCPS[0].LedNumber))
            {
                BitMatrix Qrcode = qrcodeService.Encode(dataCPS[0].LedNumber ?? "", 1);
                qrcodeService.DrawQrCode(gfx, new XPoint(25.6, 1.3), Qrcode, 0.10);
            }
            gfx.DrawString(string.Format("มหกรรมไกล่เกลี่ยชั้นบังคับคดี ครั้งที่ {0}", setdata.FestNo), font_title_bold, XBrushes.Black, new XRect(0, 1.72, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("{0}", setdata.FestName), font_title_bold, XBrushes.Black, new XRect(0, 2.4, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.FestDate ?? "")), font_title_bold, XBrushes.Black, new XRect(0, 3.1, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

            gfx.DrawString("ลำดับที่", font_normalbold, XBrushes.Black, new XRect(1.2, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].WorkNo) ? "-" : dataCPS[0].WorkNo), font_normal, XBrushes.Black, new XRect(2.3, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ลำดับกรม", font_normalbold, XBrushes.Black, new XRect(-1, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].LedNumber) ? "-" : dataCPS[0].LedNumber), font_normal, XBrushes.Black, new XRect(0.2, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

            gfx.DrawString("เลขที่คดีดำ", font_normalbold, XBrushes.Black, new XRect(21.5, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].BlackNo) ? "-" : dataCPS[0].BlackNo), font_normal, XBrushes.Black, new XRect(23.1, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ชื่อ", font_normalbold, XBrushes.Black, new XRect(1.2, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].CustomerName), font_normal, XBrushes.Black, new XRect(1.7, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString(string.Format("ภาระหนี้ ณ วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.DateAtCalulate ?? "")), font_normalbold, XBrushes.Black, new XRect(-0.3, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

            gfx.DrawString("เลขที่คดีแดง", font_normalbold, XBrushes.Black, new XRect(21.5, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].RedNo), font_normal, XBrushes.Black, new XRect(23.25, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("วันพิพากษา", font_normalbold, XBrushes.Black, new XRect(21.5, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", datehelper.doGetShortDateTHFromDBToPDF(dataCPS[0].JudgeDate ?? "")), font_normal, XBrushes.Black, new XRect(23.2, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            double diameter = 0.48; // เส้นผ่านศูนย์กลางของวงกลมในเซนติเมตร
            gfx.DrawEllipse(pen, XBrushes.White, 1.6 - diameter / 2, 6.75 - diameter / 2, diameter, diameter);
            gfx.DrawString("ตกลงรับเงื่อนไข", font_normal, XBrushes.Black, new XRect(2.1, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawEllipse(pen, XBrushes.White, 5.6 - diameter / 2, 6.75 - diameter / 2, diameter, diameter);
            gfx.DrawString("ไม่ตกลงรับเงื่อนไข  เนื่องจาก ............................................................", font_normal, XBrushes.Black, new XRect(6.1, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            #region TABLE
            #region HEADER
            XPen pentable = new XPen(XColors.Black, 0.015);
            gfx.DrawRectangle(pentable, 1.29, 7.4, 0.85, 2.1);
            gfx.DrawRectangle(pentable, 2.14, 7.4, 3.58, 2.1);
            gfx.DrawRectangle(pentable, 5.72, 7.4, 2.61, 2.1);
            gfx.DrawRectangle(pentable, 8.33, 7.4, 2.03, 2.1);
            gfx.DrawRectangle(pentable, 10.36, 7.4, 2.47, 2.1);
            gfx.DrawRectangle(pentable, 12.83, 7.4, 15.72, 2.1);
            gfx.DrawString("ที่", font_normal, XBrushes.Black, new XRect(1.62, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("หมายเลขบัตร", font_normal, XBrushes.Black, new XRect(3.0, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("พิพากษาทั้งสิ้น", font_normal, XBrushes.Black, new XRect(6.1, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("เงินต้น", font_normal, XBrushes.Black, new XRect(8.9, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ภาระหนี้", font_normal, XBrushes.Black, new XRect(11.0, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ปัจจุบัน", font_normal, XBrushes.Black, new XRect(11.1, 8.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("เงื่อนไขในการชำระ", font_normal, XBrushes.Black, new XRect(19.4, 7.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


            gfx.DrawRectangle(pentable, 12.83, 8.1, 2.7, 0.7);
            gfx.DrawRectangle(pentable, 12.83, 8.8, 2.7, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 13.3 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ปิดงวดเดียว", font_normal, XBrushes.Black, new XRect(13.75, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(13.65, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 15.53, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 15.53, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 17.70, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 16.0 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 6 งวด", font_normal, XBrushes.Black, new XRect(17.2, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(16.10, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(18.4, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 19.87, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 19.87, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 22.04, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 20.4 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 12 งวด", font_normal, XBrushes.Black, new XRect(21.5, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(20.5, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(22.8, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 24.21, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 24.21, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 26.38, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 24.8 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 24 งวด", font_normal, XBrushes.Black, new XRect(25.9, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(24.9, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(27.22, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            #endregion
            #region Detail
            double ypos_strat = 9.5;
            double ypostext_strat = 9.55;
            string dat = "-";

            for (int i = 0; i < 7; i++)
            {
                if (i != 6)
                {
                    gfx.DrawRectangle(pentable, 1.29, ypos_strat, 0.85, 0.7);
                    gfx.DrawString(string.Format("{0}", i + 1), font_normal, XBrushes.Black, new XRect(1.62, ypos_strat + 0.05, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                }
                gfx.DrawRectangle(pentable, 2.14, ypos_strat, 3.58, 0.7);
                gfx.DrawRectangle(pentable, 5.72, ypos_strat, 2.61, 0.7);
                gfx.DrawRectangle(pentable, 8.33, ypos_strat, 2.03, 0.7);
                gfx.DrawRectangle(pentable, 10.36, ypos_strat, 2.47, 0.7);
                gfx.DrawRectangle(pentable, 12.83, ypos_strat, 2.7, 0.7);
                gfx.DrawRectangle(pentable, 15.53, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 17.70, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 19.87, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 22.04, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 24.21, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 26.38, ypos_strat, 2.17, 0.7);

                if (i < dataCPS.Count) // SET VALE FROM DB
                {
                    #region calculate value
                    double judgmentAmnt = dataCPS[i].JudgmentAmnt;
                    double capitalAmnt = dataCPS[i].PrincipleAmnt; //CapitalAmnt;
                    double deptAmnt = dataCPS[i].DeptAmnt;
                    double accCloseAmnt = dataCPS[i].AccCloseAmnt;
                    double accClose6Amnt = dataCPS[i].AccClose6Amnt;
                    double installment6Amnt = dataCPS[i].Installment6Amnt;
                    double accClose12Amnt = dataCPS[i].AccClose12Amnt;
                    double installment12Amnt = dataCPS[i].Installment12Amnt;
                    double accClose24Amnt = dataCPS[i].AccClose24Amnt;
                    double installment24Amnt = dataCPS[i].Installment24Amnt;
                    string cardno = string.Format("{0}", dataCPS[i].CardNo);

                    if (dataCPS[i].Maxmonth < 6)
                    {
                        accClose6Amnt = 0;
                        installment6Amnt = 0;
                    }
                    if (dataCPS[i].Maxmonth < 12)
                    {
                        accClose12Amnt = 0;
                        installment12Amnt = 0;
                    }
                    if (dataCPS[i].Maxmonth < 24)
                    {
                        accClose24Amnt = 0;
                        installment24Amnt = 0;
                    }

                    sumjudgmentAmnt = sumjudgmentAmnt + judgmentAmnt;
                    sumcapitalAmnt = sumcapitalAmnt + capitalAmnt;
                    sumdeptAmnt = sumdeptAmnt + deptAmnt;
                    sumaccCloseAmnt = sumaccCloseAmnt + accCloseAmnt;
                    sumaccClose6Amnt = sumaccClose6Amnt + accClose6Amnt;
                    suminstallment6Amnt = suminstallment6Amnt + installment6Amnt;
                    sumaccClose12Amnt = sumaccClose12Amnt + accClose12Amnt;
                    suminstallment12Amnt = suminstallment12Amnt + installment12Amnt;
                    sumaccClose24Amnt = sumaccClose24Amnt + accClose24Amnt;
                    suminstallment24Amnt = suminstallment24Amnt + installment24Amnt;
                    #endregion

                    #region SET Value
                    if (!string.IsNullOrEmpty(cardno))
                    {
                        gfx.DrawString(cardno, font_normal, XBrushes.Black, new XRect(2.4, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(3.9, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (judgmentAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", judgmentAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-21.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (capitalAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", capitalAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-19.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (deptAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", deptAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-17.0, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accCloseAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accCloseAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-14.3, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose6Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-12.1, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment6Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-9.95, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose12Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-7.80, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment12Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-5.58, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose24Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-3.4, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment24Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-1.25, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    #endregion

                }
                else
                {
                    #region SET SUM VALUE
                    if (i == 6)
                    {
                        gfx.DrawString("รวม", font_normal_underline, XBrushes.Black, new XRect(3.69, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        if (sumjudgmentAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumjudgmentAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-21.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumcapitalAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumcapitalAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-19.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumdeptAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumdeptAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-17.0, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccCloseAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccCloseAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-14.3, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccClose6Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-12.1, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (suminstallment6Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-9.95, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccClose12Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-7.80, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (suminstallment12Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-5.58, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (sumaccClose24Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-3.4, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (suminstallment24Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-1.25, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }


                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(3.9, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    #endregion                   
                }
                ypostext_strat = ypostext_strat + 0.7;
                ypos_strat = ypos_strat + 0.7;
            }
            #region TP LP
            double y_point = 15.2;
            List<string> datashow = CreateTPDP(dataCPS);
            for (int n = 0; n < 6; n++)
            {
                gfx.DrawString(datashow[n], font_normal_small, XBrushes.Black, new XRect(12, y_point, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                y_point = y_point + 0.7;
            }
            #endregion
            #endregion
            #endregion
            gfx.DrawString("สถานะทางคดี :", font_bold_small, XBrushes.Black, new XRect(2.14, 15.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].LegalStatus), font_normal_small, XBrushes.Black, new XRect(3.8, 15.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("หมายเหตุบังคับคดี :", font_bold_small, XBrushes.Black, new XRect(2.14, 15.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].LegalExecRemark) ? "-" : dataCPS[0].LegalExecRemark), font_normal_small, XBrushes.Black, new XRect(4.3, 15.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("วันที่ยึดทรัพย์/อายัดเงินเดือน :", font_bold_small, XBrushes.Black, new XRect(2.14, 16.6, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", datehelper.doGetDateThaiFromDBToPDF(dataCPS[0].LegalExecDate ?? "-")), font_normal_small, XBrushes.Black, new XRect(5.4, 16.6, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ตรวจสอบ : _________________________________", font_normalbold, XBrushes.Black, new XRect(2.14, 18.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ผู้เจรจา : __________________________________", font_normalbold, XBrushes.Black, new XRect(18.4, 16.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString(string.Format("เจ้าหน้าที่ : | {0} | {1} | เบอร์ติดต่อ : {2}", dataCPS[0].CollectorName, string.IsNullOrEmpty(dataCPS[0].CollectorTeam) ? "-" : dataCPS[0].CollectorTeam, dataCPS[0].CollectorTel), font_normalbold, XBrushes.Black, new XRect(-3.5, 18.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);

            if (dataCPS.Count > 0)
            {
                string file_name = string.Format("Table_{0}_{1}.pdf",queueno, dataCPS[0].CustomerID);
                file_name = file_name.Replace("\n", "").Replace("\r", "").Replace("/", "").Replace(" ", "");
                string fullfilename = Path.Combine(TablePathFile, file_name);
                if (Path.Exists(TablePathFile))
                {
                    doc.Save(fullfilename);
                    return true;
                }
            }
            return false;
        }
        public void doCreteC2TableReport1File(ref PdfDocument doc,List<DataCPSCard> dataCPS, SettingData setdata,int c2count)
        {
            double spaceH = 0.6;
            string lednumber = string.Empty;
            #region C2
            doc.Info.Title = "C2 CPS DATA";
            for (int i = 0; i < dataCPS.Count; i++)
            {
                for (int c2 = 0; c2 < c2count; c2++)
                {
                    lednumber = dataCPS[i].LedNumber ?? string.Empty;
                    workno = dataCPS[0].WorkNo ?? "";
                    documentno = string.Format("ที่ RC_บค {1}_{0}", string.IsNullOrEmpty(workno) ? lednumber : workno, setdata.FestNo);

                    PdfPage page = doc.AddPage();
                    page.Size = PdfSharp.PageSize.A4;
                    page.Orientation = PdfSharp.PageOrientation.Portrait;

                    XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsUnit.Centimeter);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                    XFont font = new XFont("thsarabun", 16, XFontStyleEx.Bold, options);
                    XFont font_h_b = new XFont("thsarabun", 0.28, XFontStyleEx.Bold, options);
                    XFont font_h = new XFont("thsarabun", 0.28, XFontStyleEx.Regular, options);
                    XFont font_normalbold = new XFont("thsarabun", 0.493, XFontStyleEx.Bold, options);
                    XFont font_normal = new XFont("thsarabun", 0.493, XFontStyleEx.Regular, options);
                    XFont font_normal_small = new XFont("thsarabun", 0.35, XFontStyleEx.Regular, options);
                    XFont font_normal_install_bold = new XFont("thsarabun", 0.39, XFontStyleEx.Bold, options);
                    XFont font_normal_small_under = new XFont("thsarabun", 0.42, XFontStyleEx.Underline, options);
                    XFont font_bold_small_under = new XFont("thsarabun", 0.42, XFontStyleEx.Bold | XFontStyleEx.Underline, options);
                    XFont font_small_bold = new XFont("thsarabun", 0.35, XFontStyleEx.Bold, options);

                    #region Header
                    XImage image = XImage.FromFile(string.Format("{0}{1}", Application.StartupPath, "Images\\ktc_logo.png"));
                    gfx.DrawImage(image, 2.65, 1.25, 2.47, 1.8);
                    gfx.DrawString("บริษัท บัตรกรุงไทย จำกัด (มหาชน)", font_h_b, XBrushes.Gray, new XRect(6.17, 1.23, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("591 อาคารสมัชชาวาณิช 2 ชั้น 14 ถนนสุขุมวิท แขวงคลองตันเหนือ เขตวัฒนา กรุงเทพฯ 10110", font_h, XBrushes.Gray, new XRect(6.17, 1.52, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("โทร: 02 123 5100 โทรสาร: 02 123 5190", font_h, XBrushes.Gray, new XRect(6.17, 1.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ทะเบียนเลขที่ 0107545000110", font_h, XBrushes.Gray, new XRect(6.17, 2.08, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    gfx.DrawString("Krungthai Card Public Company Limited", font_h_b, XBrushes.Gray, new XRect(6.17, 2.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("591 United Business Center II, 14FL., Sukhumvit 33 Rd, North Klongton, Wattana, Bangkok 10110", font_h, XBrushes.Gray, new XRect(6.17, 2.75, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("Tel: 02 123 5100 Fax: 02 123 5190", font_h, XBrushes.Gray, new XRect(6.17, 3.03, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString(documentno, font_normalbold, XBrushes.Black, new XRect(2.5, 3.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.FestDate ?? "")), font_normalbold, XBrushes.Black, new XRect(-2.5, 4.63, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    #endregion
                    #region Content 1
                    gfx.DrawString("เรื่อง", font_normalbold, XBrushes.Black, new XRect(2.5, 5.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ผลการพิจารณาข้อตกลงการประนอมหนี้ตามข้อเสนอของผู้ร้อง", font_normal, XBrushes.Black, new XRect(4.0, 5.47, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("อ้างถึง", font_normalbold, XBrushes.Black, new XRect(2.5, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("บัญชีเคทีซี ของ", font_normal, XBrushes.Black, new XRect(4.0, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", dataCPS[i].CustomerName), font_normalbold, XBrushes.Black, new XRect(6.0, 5.47 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("หมายเลข", font_normal, XBrushes.Black, new XRect(4.0, 5.47 + spaceH + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(string.Format("{0}", dataCPS[i].CardNo), font_normalbold, XBrushes.Black, new XRect(5.25, 5.47 + spaceH + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                    string text1 = "ตามที่ บริษัท บัตรกรุงไทย จำกัด ( มหาชน ) หรือ “เคทีซี” ได้ยกเลิกสมาชิกพร้อมเรียกเก็บหนี้คืนทั้งหมด/";
                    XRect rect1 = new XRect(4.0, 8.11, 14.5, 10);
                    DrawJustifiedText(gfx, text1, font_normal, rect1);

                    string text2 = "ได้ดำเนินคดีกับผู้ร้อง ซึ่งศาลได้มีคำพิพากษาถึงที่สุดแล้ว และ ผู้ร้องยื่นคำร้องขอไกล่เกลี่ยข้อพิพาทในชั้นบังคับคดี ในคดีของ";
                    XRect rect2 = new XRect(2.5, 8.11 + spaceH, 16, 10);
                    DrawJustifiedText(gfx, text2, font_normal, rect2);

                    string courtname = string.Format("{0}", dataCPS[i].CourtName);
                    gfx.DrawString(courtname, font_normalbold, XBrushes.Black, new XRect(2.5, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string textredno = "หมายเลขคดีแดงที่";
                    gfx.DrawString(textredno, font_normal, XBrushes.Black, new XRect(gfx.MeasureString(courtname, font_normalbold).Width + 2.8, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


                    double widthdata = gfx.MeasureString(textredno, font_normal).Width + gfx.MeasureString(courtname, font_normalbold).Width;
                    string rednumber = string.Format("{0}", dataCPS[i].RedNo);
                    gfx.DrawString(rednumber, font_normalbold, XBrushes.Black, new XRect(widthdata + 3, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string textktc = "ระหว่าง บริษัท บัตรกรุงไทย จำกัด ( มหาชน ) โจทก์ กับ";
                    gfx.DrawString(textktc, font_normal, XBrushes.Black, new XRect(-2.5, 8.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);

                    gfx.DrawString(string.Format("{0} จำเลย", dataCPS[i].CustomerName), font_normal, XBrushes.Black, new XRect(2.5, 8.9 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string text3 = "ต่อมาผู้ร้องได้ขอลดหย่อนภาระหนี้ เพื่อชำระหนี้ให้กับบริษัทฯ เป็นการเสร็จสิ้นตามความสามารถการชำระหนี้";
                    XRect rect3 = new XRect(4.0, 10.5, 14.5, 10);
                    DrawJustifiedText(gfx, text3, font_normal, rect3);

                    gfx.DrawString("ของผู้ร้องและบริษัทฯ ได้อนุมัติลดหย่อนภาระหนี้ โดยมีเงื่อนไขให้รับชำระหนี้เสร็จสิ้น ดังนี้", font_normal, XBrushes.Black, new XRect(2.5, 10.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    #endregion
                    #region TABLE

                    XBrush backgroundBrush = XBrushes.LightGray;
                    XBrush whiteBrush = XBrushes.Gainsboro;
                    //Table HD
                    XPen pen = new XPen(XColors.Black, 0.01);
                    gfx.DrawRectangle(pen, backgroundBrush, 2.2, 11.50, 4.28, 1.76);
                    gfx.DrawRectangle(pen, backgroundBrush, 6.48, 11.50, 2.7, 1.76);
                    gfx.DrawRectangle(pen, backgroundBrush, 9.18, 11.50, 2.9, 1.76);

                    gfx.DrawRectangle(pen, backgroundBrush, 12.08, 11.50, 2.20, 1.13);
                    gfx.DrawRectangle(pen, backgroundBrush, 14.28, 11.50, 2.30, 1.13);
                    gfx.DrawRectangle(pen, backgroundBrush, 16.48, 11.50, 2.30, 1.13);

                    gfx.DrawRectangle(pen, backgroundBrush, 12.08, 12.63, 2.20, 0.63);
                    gfx.DrawRectangle(pen, backgroundBrush, 14.28, 12.63, 2.30, 0.63);
                    gfx.DrawRectangle(pen, backgroundBrush, 16.48, 12.63, 2.30, 0.63);

                    ////Detail

                    gfx.DrawRectangle(pen, 2.2, 13.26, 4.28, 0.63);
                    gfx.DrawRectangle(pen, 2.2, 13.89, 4.28, 0.63);
                    gfx.DrawRectangle(pen, 2.2, 14.52, 4.28, 0.63);
                    gfx.DrawRectangle(pen, 2.2, 15.15, 4.28, 0.63);
                    gfx.DrawRectangle(pen, 2.2, 15.78, 4.28, 0.63);

                    gfx.DrawRectangle(pen, 6.48, 13.26, 2.7, 0.63);
                    gfx.DrawRectangle(pen, 6.48, 13.89, 2.7, 0.63);
                    gfx.DrawRectangle(pen, 6.48, 14.52, 2.7, 0.63);
                    gfx.DrawRectangle(pen, 6.48, 15.15, 2.7, 0.63);
                    gfx.DrawRectangle(pen, 6.48, 15.78, 2.7, 0.63);

                    gfx.DrawRectangle(pen, whiteBrush, 9.18, 13.26, 2.9, 0.63);
                    gfx.DrawRectangle(pen, 9.18, 13.89, 2.9, 0.63);
                    gfx.DrawRectangle(pen, 9.18, 14.52, 2.9, 0.63);
                    gfx.DrawRectangle(pen, 9.18, 15.15, 2.9, 0.63);
                    gfx.DrawRectangle(pen, 9.18, 15.78, 2.9, 0.63);


                    gfx.DrawRectangle(pen, whiteBrush, 12.08, 13.26, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 12.08, 13.89, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 12.08, 14.52, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 12.08, 15.15, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 12.08, 15.78, 2.20, 0.63);

                    gfx.DrawRectangle(pen, whiteBrush, 14.28, 13.26, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 14.28, 13.89, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 14.28, 14.52, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 14.28, 15.15, 2.20, 0.63);
                    gfx.DrawRectangle(pen, 14.28, 15.78, 2.20, 0.63);

                    gfx.DrawRectangle(pen, whiteBrush, 16.48, 13.26, 2.30, 0.63);
                    gfx.DrawRectangle(pen, 16.48, 13.89, 2.30, 0.63);
                    gfx.DrawRectangle(pen, 16.48, 14.52, 2.30, 0.63);
                    gfx.DrawRectangle(pen, 16.48, 15.15, 2.30, 0.63);
                    gfx.DrawRectangle(pen, 16.48, 15.78, 2.30, 0.63);

                    gfx.DrawString("เงื่อนไขการชำระ", font_normal_install_bold, XBrushes.Black, new XRect(3.5, 12.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("จำนวนเงินที่ชำระ", font_normal_install_bold, XBrushes.Black, new XRect(6.85, 12.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("จำนวนเงินที่ชำระ", font_normal_install_bold, XBrushes.Black, new XRect(9.7, 12.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ต่องวด", font_normal_install_bold, XBrushes.Black, new XRect(10.3, 12.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(12.85, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(12.4, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(12.45, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(15.05, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(14.6, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(14.65, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("งวดที่", font_normal_install_bold, XBrushes.Black, new XRect(17.25, 11.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("..........ถึง..........", font_normal_install_bold, XBrushes.Black, new XRect(16.8, 12.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("งวดละ (บาท)", font_normal_install_bold, XBrushes.Black, new XRect(16.85, 12.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    double diameter = 0.35; // เส้นผ่านศูนย์กลางของวงกลมในเซนติเมตร
                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 13.56 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ชำระปิดบัญชีคราวเดียว", font_normal_small, XBrushes.Black, new XRect(2.85, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(10.6, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(13.1, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(15.3, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("-", font_normal_install_bold, XBrushes.Black, new XRect(17.6, 13.35, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 14.19 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อนชำระไม่เกิน          6          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 13.98, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 14.82 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อนชำระไม่เกิน        12          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 14.61, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 15.45 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อนชำระไม่เกิน        24          งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 15.24, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);

                    gfx.DrawEllipse(pen, XBrushes.White, 2.5 - diameter / 2, 16.08 - diameter / 2, diameter, diameter);
                    gfx.DrawString("ผ่อนชำระไม่เกิน ......................... งวด", font_normal_small, XBrushes.Black, new XRect(2.85, 15.87, page.Width.Point, page.Height.Point), XStringFormats.TopLeft);



                    #endregion
                    #region Content 2
                    gfx.DrawString("โดยเริ่มชำระงวดแรกภายในวันที่ .................................. งวดต่อไปชำระ ทุกๆ วันที่ ...................... ของทุกเดือน โดย", font_normal, XBrushes.Black, new XRect(4.0, 16.65, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ชำระให้เสร็จสิ้นภายในวันที่ ...............................................ทั้งนี้บริษัทฯ อาจจะมีการแจ้งเตือนก่อนวันถึงกำหนดชำระผ่าน SMS", font_normal, XBrushes.Black, new XRect(2.5, 16.65 + spaceH, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string custtel = string.Format("{0}", dataCPS[i].CustomerTel);
                    string textcusttel = "ตามหมายเลขโทรศัพท์ที่ท่านได้ให้ไว้แก่บริษัทฯ ครั้งหลังสุด";
                    string textcusttelother = "หรือ ...................................................";
                    gfx.DrawString(textcusttel, font_normal, XBrushes.Black, new XRect(2.5, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(custtel, font_normalbold, XBrushes.Black, new XRect(gfx.MeasureString(textcusttel, font_normal).Width + 2.65, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    double withcustel = gfx.MeasureString(textcusttel, font_normal).Width + gfx.MeasureString(custtel, font_normalbold).Width;
                    gfx.DrawString(textcusttelother, font_normal, XBrushes.Black, new XRect(withcustel + 2.8, 17.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    string text4 = "ภายในกำหนดข้างต้นหากผู้ร้องชำระหนี้ดังกล่าวให้กับบริษัทฯเป็นที่ครบถ้วนเรียบร้อยแล้ว บริษัทฯจะถือการชำระ";
                    XRect rect4 = new XRect(4.0, 18.90, 14.5, 10);
                    DrawJustifiedText(gfx, text4, font_normal, rect4);

                    string text5 = "หนี้ตามจำนวนดังกล่าวเป็นการชำระหนี้เสร็จสิ้นตามที่บริษัทฯอนุมัติ โดยบริษัทฯจักได้ดำเนินการปรับปรุงบัญชีของผู้ร้องต่อไป";
                    XRect rect5 = new XRect(2.5, 18.90 + spaceH, 16, 10);
                    DrawJustifiedText(gfx, text5, font_normal, rect5);

                    string txtcontact = "ติดต่อพนักงานผู้รับผิดชอบ ";
                    string txtcollecttor = string.Format("{0} ", dataCPS[i].CollectorName);
                    string txttel = "โทร. ";
                    string txtcolltel = string.Format("{0}", dataCPS[i].CollectorTel);
                    gfx.DrawString(txtcontact, font_normal_small_under, XBrushes.Black, new XRect(2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(txtcollecttor, font_bold_small_under, XBrushes.Black, new XRect(gfx.MeasureString(txtcontact, font_normal_small_under).Width + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    double withcontact = gfx.MeasureString(txtcontact, font_normal_small_under).Width + gfx.MeasureString(txtcollecttor, font_bold_small_under).Width;
                    gfx.DrawString(txttel, font_normal_small_under, XBrushes.Black, new XRect(withcontact + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString(txtcolltel, font_bold_small_under, XBrushes.Black, new XRect(withcontact + gfx.MeasureString(txttel, font_normal_small_under).Width + 2.5, 19.73, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("หมายเหตุ", font_normal_small, XBrushes.Black, new XRect(2.5, 21.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("1. กรณีมีการชำระหนี้ข้างต้นถือเป็นการชำระหนี้เสร็จสิ้นเฉพาะบัตรเลขที่ดังกล่าวข้างต้นเท่านั้น ไม่รวมถึงภาระหนี้อื่นของ", font_normal_small, XBrushes.Black, new XRect(4.0, 21.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("บมจ.บัตรกรุงไทย (ถ้ามี)", font_normal_small, XBrushes.Black, new XRect(4.0, 21.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("2. หากผู้ร้องผิดนัด ไม่ชำระหนี้ตามข้อตกลงไม่ว่าข้อหนึ่งข้อใด ถือว่าข้อตกลงไหล่เกลี่ยเป็นอันยกเลิก และบมจ.บัตรกรุงไทยจะดำเนินการ", font_normal_small, XBrushes.Black, new XRect(4.0, 22.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("บังคับคดีตามกฎหมายต่อไป โดยกลับไปคิดยอดหนี้ตามคำพิพากษาของศาลส่วนเงินที่ผู้ร้องชำระเข้ามาตามข้อตกลงไกล่เกลี่ย ให้เป็นส่วน", font_normal_small, XBrushes.Black, new XRect(4.0, 22.8, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("หนึ่งของการชำระหนี้ตามคำพิพากษา (ถ้ามี)", font_normal_small, XBrushes.Black, new XRect(4.0, 23.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

                    gfx.DrawString("ลงชื่อ .......................................................... ผู้ร้องขอไกล่เกลี่ย", font_normal, XBrushes.Black, new XRect(2.5, 25.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("ลงชื่อ .......................................................... ผู้แทนโจทก์", font_normal, XBrushes.Black, new XRect(-2.5, 25.0, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    gfx.DrawString("(                                              )", font_normal, XBrushes.Black, new XRect(3.1, 25.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    gfx.DrawString("(                                              )", font_normal, XBrushes.Black, new XRect(-3.85, 25.7, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    gfx.DrawString(string.Format("ลำดับกรม {0}", dataCPS[i].LedNumber), font_small_bold, XBrushes.Black, new XRect(-2.5, 27.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    #endregion
                }
            }
               doCreateTable(dataCPS, setdata,ref doc);
            #endregion            
        }
        private void doCreateTable(List<DataCPSCard> dataCPS, SettingData setdata,ref PdfDocument doc)
        {
            #region Table
           
            double sumjudgmentAmnt = 0;
            double sumcapitalAmnt = 0;
            double sumdeptAmnt = 0;
            double sumaccCloseAmnt = 0;
            double sumaccClose6Amnt = 0;
            double suminstallment6Amnt = 0;
            double sumaccClose12Amnt = 0;
            double suminstallment12Amnt = 0;
            double sumaccClose24Amnt = 0;
            double suminstallment24Amnt = 0;

            PdfPage page = doc.AddPage();
            page.Size = PdfSharp.PageSize.A4;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsUnit.Centimeter);
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            XFont font_title_bold = new XFont("thsarabun", 0.56, XFontStyleEx.Bold, options);
            XFont font_normalbold = new XFont("thsarabun", 0.49, XFontStyleEx.Bold, options);
            XFont font_normal = new XFont("thsarabun", 0.49, XFontStyleEx.Regular, options);
            XFont font_normal_underline = new XFont("thsarabun", 0.49, XFontStyleEx.Underline, options);
            XFont font_normal_small = new XFont("thsarabun", 0.38, XFontStyleEx.Regular, options);
            XFont font_bold_small = new XFont("thsarabun", 0.38, XFontStyleEx.Bold, options);
            XPen pen = new XPen(XColors.Black, 0.03);

            XImage image = XImage.FromFile(string.Format("{0}{1}", Application.StartupPath, @"Images/ktc_logo.png"));
            gfx.DrawImage(image, 1.1, 1.72, 3.24, 2.33);

            if (!string.IsNullOrEmpty(dataCPS[0].LedNumber))
            {
                BitMatrix Qrcode = qrcodeService.Encode(dataCPS[0].LedNumber ?? "", 1);
                qrcodeService.DrawQrCode(gfx, new XPoint(25.6, 1.3), Qrcode, 0.10);
            }
            gfx.DrawString(string.Format("มหกรรมไกล่เกลี่ยชั้นบังคับคดี ครั้งที่ {0}", setdata.FestNo), font_title_bold, XBrushes.Black, new XRect(0, 1.72, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("{0}", setdata.FestName), font_title_bold, XBrushes.Black, new XRect(0, 2.4, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.FestDate ?? "")), font_title_bold, XBrushes.Black, new XRect(0, 3.1, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

            gfx.DrawString("ลำดับที่", font_normalbold, XBrushes.Black, new XRect(1.2, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].WorkNo) ? "-" : dataCPS[0].WorkNo), font_normal, XBrushes.Black, new XRect(2.3, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ลำดับกรม", font_normalbold, XBrushes.Black, new XRect(-1, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].LedNumber) ? "-" : dataCPS[0].LedNumber), font_normal, XBrushes.Black, new XRect(0.2, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

            gfx.DrawString("เลขที่คดีดำ", font_normalbold, XBrushes.Black, new XRect(21.5, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].BlackNo) ? "-" : dataCPS[0].BlackNo), font_normal, XBrushes.Black, new XRect(23.1, 4.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ชื่อ", font_normalbold, XBrushes.Black, new XRect(1.2, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].CustomerName), font_normal, XBrushes.Black, new XRect(1.7, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString(string.Format("ภาระหนี้ ณ วันที่ {0}", datehelper.doGetDateThaiFromDBToPDF(setdata.DateAtCalulate ?? "")), font_normalbold, XBrushes.Black, new XRect(-0.3, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopCenter);

            gfx.DrawString("เลขที่คดีแดง", font_normalbold, XBrushes.Black, new XRect(21.5, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].RedNo), font_normal, XBrushes.Black, new XRect(23.25, 5.55, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("วันพิพากษา", font_normalbold, XBrushes.Black, new XRect(21.5, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", datehelper.doGetShortDateTHFromDBToPDF(dataCPS[0].JudgeDate ?? "")), font_normal, XBrushes.Black, new XRect(23.2, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            double diameter = 0.48; // เส้นผ่านศูนย์กลางของวงกลมในเซนติเมตร
            gfx.DrawEllipse(pen, XBrushes.White, 1.6 - diameter / 2, 6.75 - diameter / 2, diameter, diameter);
            gfx.DrawString("ตกลงรับเงื่อนไข", font_normal, XBrushes.Black, new XRect(2.1, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawEllipse(pen, XBrushes.White, 5.6 - diameter / 2, 6.75 - diameter / 2, diameter, diameter);
            gfx.DrawString("ไม่ตกลงรับเงื่อนไข  เนื่องจาก ............................................................", font_normal, XBrushes.Black, new XRect(6.1, 6.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            #region TABLE
            #region HEADER
            XPen pentable = new XPen(XColors.Black, 0.015);
            gfx.DrawRectangle(pentable, 1.29, 7.4, 0.85, 2.1);
            gfx.DrawRectangle(pentable, 2.14, 7.4, 3.58, 2.1);
            gfx.DrawRectangle(pentable, 5.72, 7.4, 2.61, 2.1);
            gfx.DrawRectangle(pentable, 8.33, 7.4, 2.03, 2.1);
            gfx.DrawRectangle(pentable, 10.36, 7.4, 2.47, 2.1);
            gfx.DrawRectangle(pentable, 12.83, 7.4, 15.72, 2.1);
            gfx.DrawString("ที่", font_normal, XBrushes.Black, new XRect(1.62, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("หมายเลขบัตร", font_normal, XBrushes.Black, new XRect(3.0, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("พิพากษาทั้งสิ้น", font_normal, XBrushes.Black, new XRect(6.1, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("เงินต้น", font_normal, XBrushes.Black, new XRect(8.9, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ภาระหนี้", font_normal, XBrushes.Black, new XRect(11.0, 7.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ปัจจุบัน", font_normal, XBrushes.Black, new XRect(11.1, 8.5, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("เงื่อนไขในการชำระ", font_normal, XBrushes.Black, new XRect(19.4, 7.45, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);


            gfx.DrawRectangle(pentable, 12.83, 8.1, 2.7, 0.7);
            gfx.DrawRectangle(pentable, 12.83, 8.8, 2.7, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 13.3 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ปิดงวดเดียว", font_normal, XBrushes.Black, new XRect(13.75, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(13.65, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 15.53, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 15.53, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 17.70, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 16.0 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 6 งวด", font_normal, XBrushes.Black, new XRect(17.2, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(16.10, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(18.4, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 19.87, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 19.87, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 22.04, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 20.4 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 12 งวด", font_normal, XBrushes.Black, new XRect(21.5, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(20.5, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(22.8, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawRectangle(pentable, 24.21, 8.1, 4.34, 0.7);
            gfx.DrawRectangle(pentable, 24.21, 8.8, 2.17, 0.7);
            gfx.DrawRectangle(pentable, 26.38, 8.8, 2.17, 0.7);
            gfx.DrawEllipse(pen, XBrushes.White, 24.8 - diameter / 2, 8.45 - diameter / 2, diameter, diameter);
            gfx.DrawString("ผ่อน 24 งวด", font_normal, XBrushes.Black, new XRect(25.9, 8.15, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("ยอดชำระ", font_normal, XBrushes.Black, new XRect(24.9, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString("งวดละ", font_normal, XBrushes.Black, new XRect(27.22, 8.85, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            #endregion
            #region Detail
            double ypos_strat = 9.5;
            double ypostext_strat = 9.55;
            string dat = "-";

            for (int i = 0; i < 7; i++)
            {
                if (i != 6)
                {
                    gfx.DrawRectangle(pentable, 1.29, ypos_strat, 0.85, 0.7);
                    gfx.DrawString(string.Format("{0}", i + 1), font_normal, XBrushes.Black, new XRect(1.62, ypos_strat + 0.05, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                }
                gfx.DrawRectangle(pentable, 2.14, ypos_strat, 3.58, 0.7);
                gfx.DrawRectangle(pentable, 5.72, ypos_strat, 2.61, 0.7);
                gfx.DrawRectangle(pentable, 8.33, ypos_strat, 2.03, 0.7);
                gfx.DrawRectangle(pentable, 10.36, ypos_strat, 2.47, 0.7);
                gfx.DrawRectangle(pentable, 12.83, ypos_strat, 2.7, 0.7);
                gfx.DrawRectangle(pentable, 15.53, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 17.70, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 19.87, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 22.04, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 24.21, ypos_strat, 2.17, 0.7);
                gfx.DrawRectangle(pentable, 26.38, ypos_strat, 2.17, 0.7);

                if (i < dataCPS.Count) // SET VALE FROM DB
                {
                    #region calculate value
                    double judgmentAmnt = dataCPS[i].JudgmentAmnt;
                    double capitalAmnt = dataCPS[i].PrincipleAmnt; //CapitalAmnt;
                    double deptAmnt = dataCPS[i].DeptAmnt;
                    double accCloseAmnt = dataCPS[i].AccCloseAmnt;
                    double accClose6Amnt = dataCPS[i].AccClose6Amnt;
                    double installment6Amnt = dataCPS[i].Installment6Amnt;
                    double accClose12Amnt = dataCPS[i].AccClose12Amnt;
                    double installment12Amnt = dataCPS[i].Installment12Amnt;
                    double accClose24Amnt = dataCPS[i].AccClose24Amnt;
                    double installment24Amnt = dataCPS[i].Installment24Amnt;
                    string cardno = string.Format("{0}", dataCPS[i].CardNo);

                    if (dataCPS[i].Maxmonth < 6)
                    {
                        accClose6Amnt = 0;
                        installment6Amnt = 0;
                    }
                    if (dataCPS[i].Maxmonth < 12)
                    {
                        accClose12Amnt = 0;
                        installment12Amnt = 0;
                    }
                    if (dataCPS[i].Maxmonth < 24)
                    {
                        accClose24Amnt = 0;
                        installment24Amnt = 0;
                    }

                    sumjudgmentAmnt = sumjudgmentAmnt + judgmentAmnt;
                    sumcapitalAmnt = sumcapitalAmnt + capitalAmnt;
                    sumdeptAmnt = sumdeptAmnt + deptAmnt;
                    sumaccCloseAmnt = sumaccCloseAmnt + accCloseAmnt;
                    sumaccClose6Amnt = sumaccClose6Amnt + accClose6Amnt;
                    suminstallment6Amnt = suminstallment6Amnt + installment6Amnt;
                    sumaccClose12Amnt = sumaccClose12Amnt + accClose12Amnt;
                    suminstallment12Amnt = suminstallment12Amnt + installment12Amnt;
                    sumaccClose24Amnt = sumaccClose24Amnt + accClose24Amnt;
                    suminstallment24Amnt = suminstallment24Amnt + installment24Amnt;
                    #endregion

                    #region SET Value
                    if (!string.IsNullOrEmpty(cardno))
                    {
                        gfx.DrawString(cardno, font_normal, XBrushes.Black, new XRect(2.4, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(3.9, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (judgmentAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", judgmentAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-21.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (capitalAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", capitalAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-19.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (deptAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", deptAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-17.0, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accCloseAmnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accCloseAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-14.3, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose6Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-12.1, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment6Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-9.95, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose12Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-7.80, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment12Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-5.58, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (accClose24Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", accClose24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-3.4, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }

                    if (installment24Amnt > 0)
                    {
                        gfx.DrawString(string.Format("{0}", installment24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-1.25, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    #endregion

                }
                else
                {
                    #region SET SUM VALUE
                    if (i == 6)
                    {
                        gfx.DrawString("รวม", font_normal_underline, XBrushes.Black, new XRect(3.69, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        if (sumjudgmentAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumjudgmentAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-21.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumcapitalAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumcapitalAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-19.47, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumdeptAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumdeptAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-17.0, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccCloseAmnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccCloseAmnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-14.3, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccClose6Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-12.1, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (suminstallment6Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment6Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-9.95, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }

                        if (sumaccClose12Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-7.80, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (suminstallment12Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment12Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-5.58, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (sumaccClose24Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", sumaccClose24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-3.4, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }
                        if (suminstallment24Amnt > 0)
                        {
                            gfx.DrawString(string.Format("{0}", suminstallment24Amnt.ToString("N2")), font_normal, XBrushes.Black, new XRect(-1.25, ypos_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);
                        }
                        else
                        {
                            gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        }


                    }
                    else
                    {
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(3.9, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(7.03, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(9.35, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(11.6, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(14.18, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(16.62, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(18.79, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(20.96, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(23.13, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(25.3, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                        gfx.DrawString(dat, font_normal, XBrushes.Black, new XRect(27.47, ypostext_strat, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                    }
                    #endregion                   
                }
                ypostext_strat = ypostext_strat + 0.7;
                ypos_strat = ypos_strat + 0.7;
            }
            #region TP LP
            double y_point = 15.2;
            List<string> datashow = CreateTPDP(dataCPS);
            for (int n = 0; n < 6; n++)
            {
                gfx.DrawString(datashow[n], font_normal_small, XBrushes.Black, new XRect(12, y_point, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
                y_point = y_point + 0.7;
            }
            #endregion
            #endregion
            #endregion
            gfx.DrawString("สถานะทางคดี :", font_bold_small, XBrushes.Black, new XRect(2.14, 15.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", dataCPS[0].LegalStatus), font_normal_small, XBrushes.Black, new XRect(3.8, 15.2, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("หมายเหตุบังคับคดี :", font_bold_small, XBrushes.Black, new XRect(2.14, 15.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", string.IsNullOrEmpty(dataCPS[0].LegalExecRemark) ? "-" : dataCPS[0].LegalExecRemark), font_normal_small, XBrushes.Black, new XRect(4.3, 15.9, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("วันที่ยึดทรัพย์/อายัดเงินเดือน :", font_bold_small, XBrushes.Black, new XRect(2.14, 16.6, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);
            gfx.DrawString(string.Format("{0}", datehelper.doGetDateThaiFromDBToPDF(dataCPS[0].LegalExecDate ?? "-")), font_normal_small, XBrushes.Black, new XRect(5.4, 16.6, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ตรวจสอบ : _________________________________", font_normalbold, XBrushes.Black, new XRect(2.14, 18.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString("ผู้เจรจา : __________________________________", font_normalbold, XBrushes.Black, new XRect(18.4, 16.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopLeft);

            gfx.DrawString(string.Format("เจ้าหน้าที่ : | {0} | {1} | เบอร์ติดต่อ : {2}", dataCPS[0].CollectorName, string.IsNullOrEmpty(dataCPS[0].CollectorTeam) ? "-" : dataCPS[0].CollectorTeam, dataCPS[0].CollectorTel), font_normalbold, XBrushes.Black, new XRect(-3.5, 18.3, page.Width.Centimeter, page.Height.Centimeter), XStringFormats.TopRight);

            #endregion           
        }
        #endregion
        #region Other Method
        private List<string> CreateTPDP(List<DataCPSCard> dataCPS)
        {
            List<string> DPTPtextshow =  new List<string>();
            string text_ = string.Empty;
            for(int i = 0; i < 6; i++)
            {
                if(i < dataCPS.Count)
                {
                    string datetp = datehelper.doGetShortDateTHFromDBToPDF(dataCPS[i].LastPayDate??"-");
                    double tppay = dataCPS[i].PayAfterJudgAmt;
                    text_ = string.Format("TP{0} : {1} | LPD : {2}",i+1, tppay == 0?"-":tppay.ToString("N2"), datetp);
                }
                else
                {
                    text_ =  string.Format("TP{0} : - | LPD : -", i + 1, i + 1);
                }
                DPTPtextshow.Add(text_);
            }
            return DPTPtextshow;
        }
        private void DrawJustifiedText(XGraphics gfx, string text, XFont font, XRect rect)
        {
            // แบ่งข้อความเป็นคำ
            string[] words = text.Split(' ');
            double lineSpacing = gfx.MeasureString("A", font).Height;

            List<string> lines = new List<string>();
            string line = "";
            foreach (string word in words)
            {
                string testLine = (line.Length == 0) ? word : line + " " + word;
                XSize size = gfx.MeasureString(testLine, font);
                if (size.Width > rect.Width)
                {
                    lines.Add(line);
                    line = word;
                }
                else
                {
                    line = testLine;
                }
            }
            lines.Add(line);

            double y = rect.Top;
            foreach (string l in lines)
            {
                DrawJustifiedLine(gfx, l, font, new XPoint(rect.Left, y), rect.Width);
                y += lineSpacing;
            }
        }
        private void DrawJustifiedLine(XGraphics gfx, string line, XFont font, XPoint pt, double width)
        {
            string[] words = line.Split(' ');
            double spaceWidth = gfx.MeasureString(" ", font).Width;
            double totalWidth = gfx.MeasureString(line, font).Width;

            if (words.Length > 1)
            {
                double gap = (width - totalWidth) / (words.Length - 1);
                double x = pt.X;
                foreach (string word in words)
                {
                    gfx.DrawString(word, font, XBrushes.Black, x, pt.Y);
                    x += gfx.MeasureString(word, font).Width + spaceWidth + gap;
                }
            }
            else
            {
                gfx.DrawString(line, font, XBrushes.Black, pt);
            }
        }
        #endregion
    }
}
