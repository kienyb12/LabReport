using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Lab1.Models;
using Lab1.Models.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Lab1.Controllers
{
    public class ReportController : Controller
    {
        // GET: Report
        public ActionResult Index()
        {
            using (var db = new Hungtri2019Entities())
            {
                var date = DateTime.Today.AddDays(-1);
                var items = db.TblHistoryErrors.Where(x => x.CreateDate.Value >= date && x.Val != "").Take(10).OrderByDescending(x => x.CreateDate).ToList();
                return View(items);
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public FileResult Export(string GridHtml)
        {
            using (MemoryStream stream = new System.IO.MemoryStream())
            {
                StringReader sr = new StringReader(GridHtml);
                Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 100f, 0f);
                PdfWriter writer = PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                XMLWorkerHelper.GetInstance().ParseXHtml(writer, pdfDoc, sr);
                pdfDoc.Close();
                return File(stream.ToArray(), "application/pdf", "Grid.pdf");
            }
        }

        [HttpPost]
        [ActionName("Index_Post")]
        public ActionResult Index_Post()
        {
            Document pdfDoc = new Document(PageSize.A4, 25, 25, 25, 15);
            PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
            pdfWriter.PageEvent = new ITextEvents(Server.MapPath("~/Content/Upload/salma.jpg"));
            pdfDoc.Open();

            //Top Heading
            Chunk chunk;//= new Chunk("Your Credit Card Statement Report has been Generated", FontFactory.GetFont("Arial", 20, Font.BOLDITALIC, BaseColor.MAGENTA));
            //chunk = new Chunk("Your Credit Card Statement Report has been Generated", FontFactory.GetFont("Arial", 20, Font.BOLDITALIC, BaseColor.MAGENTA));
            //pdfDoc.Add(chunk);

            //Horizontal Line
            Paragraph line;//= new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
            //line= new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
            //pdfDoc.Add(line);

            //Table
            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 100;
            //0=Left, 1=Centre, 2=Right
            table.HorizontalAlignment = 0;
            table.SpacingBefore = 20f;
            table.SpacingAfter = 30f;

            //Cell no 1
            PdfPCell cell = new PdfPCell();
            cell.Border = 0;
            Image image = Image.GetInstance(Server.MapPath("~/Content/Upload/salma.jpg"));
            image.ScaleAbsolute(200, 150);
            cell.AddElement(image);
            table.AddCell(cell);

            //Cell no 2
            chunk = new Chunk("Name: Mrs. Salma Mukherji,\nAddress: Latham Village, Latham, New York, US, \nOccupation: Nurse, \nAge: 35 years", FontFactory.GetFont("Arial", 15, Font.NORMAL, BaseColor.PINK));
            cell = new PdfPCell();
            cell.Border = 0;
            cell.AddElement(chunk);
            table.AddCell(cell);

            //Add table to document
            pdfDoc.Add(table);

            //Horizontal Line
            line = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
            pdfDoc.Add(line);

            //Table
            table = new PdfPTable(5);
            table.WidthPercentage = 100;
            table.HorizontalAlignment = 0;
            table.SpacingBefore = 20f;
            table.SpacingAfter = 30f;

            //Cell
            cell = new PdfPCell();
            chunk = new Chunk("This Month's Transactions of your Credit Card");
            cell.AddElement(chunk);
            cell.Colspan = 5;
            cell.BackgroundColor = BaseColor.PINK;
            table.AddCell(cell);

            table.AddCell("S.No");
            table.AddCell("NYC Junction");
            table.AddCell("Item");
            table.AddCell("Cost");
            table.AddCell("Date");

            table.AddCell("1");
            table.AddCell("David Food Store");
            table.AddCell("Fruits & Vegetables");
            table.AddCell("$100.00");
            table.AddCell("June 1");

            table.AddCell("2");
            table.AddCell("Child Store");
            table.AddCell("Diaper Pack");
            table.AddCell("$6.00");
            table.AddCell("June 9");

            table.AddCell("3");
            table.AddCell("Punjabi Restaurant");
            table.AddCell("Dinner");
            table.AddCell("$29.00");
            table.AddCell("June 15");

            table.AddCell("4");
            table.AddCell("Wallmart Albany");
            table.AddCell("Grocery");
            table.AddCell("$299.50");
            table.AddCell("June 25");

            table.AddCell("5");
            table.AddCell("Singh Drugs");
            table.AddCell("Back Pain Tablets");
            table.AddCell("$14.99");
            table.AddCell("June 28");

            table.AddCell("6");
            table.AddCell("Singh Drugs");
            table.AddCell("Back Pain Tablets");
            table.AddCell("$14.99");
            table.AddCell("June 28");

            table.AddCell("7");
            table.AddCell("Singh Drugs");
            table.AddCell("Back Pain Tablets");
            table.AddCell("$14.99");
            table.AddCell("June 28");

            pdfDoc.Add(table);

            Paragraph para = new Paragraph();
            para.Add("Hello Salma,\n\nThank you for being our valuable customer. We hope our letter finds you in the best of health and wealth.\n\nYours Sincerely, \nBank of America");
            pdfDoc.Add(para);

            //Horizontal Line
            line = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
            pdfDoc.Add(line);

            para = new Paragraph();
            para.Add("This PDF is generated using iTextSharp. You can read the turorial:");
            para.SpacingBefore = 20f;
            para.SpacingAfter = 20f;
            pdfDoc.Add(para);

            //Creating link
            chunk = new Chunk("How to Create a Pdf File");
            chunk.Font = FontFactory.GetFont("Arial", 25, Font.BOLD, BaseColor.RED);
            chunk.SetAnchor("https://www.yogihosting.com/create-pdf-asp-net-mvc/");
            pdfDoc.Add(chunk);

            pdfWriter.CloseStream = false;
            pdfDoc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=Credit-Card-Report.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.Write(pdfDoc);
            Response.End();

            return View();
        }

        [HttpPost]
        [ActionName("TmpReport")]
        public ActionResult TmpReport()
        {
            Document pdfDoc = new Document(PageSize.A4, 25, 25, 140f, 15);
            PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
            pdfWriter.PageEvent = new ITextEvents(Server.MapPath("~/Content/Upload/logo.png"));
            pdfDoc.Open();

            //for (int i = 0; i < 10; i++)
            //{
            //    Paragraph para = new Paragraph("Hello world. Checking Header Footer", new Font(Font.FontFamily.HELVETICA, 22));
            //    para.Alignment = Element.ALIGN_CENTER;
            //    pdfDoc.Add(para);
            //    pdfDoc.NewPage();
            //}
            //Chunk chunk = new Chunk("Your Credit Card Statement Report has been Generated", FontFactory.GetFont("Arial", 20, Font.BOLDITALIC, BaseColor.MAGENTA));
            //pdfDoc.Add(chunk);
            PdfPTable table = new PdfPTable(2);
            table.HeaderRows = 1;
            table.WidthPercentage = 100;
            table.HorizontalAlignment = 0;
            table.SpacingBefore = 20f;
            table.SpacingAfter = 30f;
            PdfPCell cell;
            Chunk chunk;
            cell = new PdfPCell();
            chunk = new Chunk("Bao Cao Tang soi");
            cell.AddElement(chunk);
            cell.Colspan = 2;
            cell.BackgroundColor = BaseColor.PINK;
            table.AddCell(cell);

            table.AddCell(new PdfPCell(new Phrase("Thoi gian Chay")) { BackgroundColor = BaseColor.GREEN });
            table.AddCell(new PdfPCell(new Phrase("Danh sach loi")) { BackgroundColor = BaseColor.GREEN });
            using (var db = new Hungtri2019Entities())
            {
                var date = DateTime.Today.AddDays(-2);
                var items = db.TblHistoryErrors.Where(x => x.CreateDate.Value >= date && x.Val != "").OrderByDescending(x => x.CreateDate).Take(20).ToList();
                var n = 0;
                foreach (var item in items)
                {
                    table.AddCell(item.CreateDate.Value.ToString("MM/dd/yyyy H:mm:ss"));
                    PdfPTable pdfPTable = new PdfPTable(2);
                    foreach (var val in item.Val.Vals())
                    {
                        n += 1;
                        pdfPTable.AddCell(new PdfPCell(new Phrase(string.Format("{0}.... {1}", n, val.Key))) { Border = 0 });
                        pdfPTable.AddCell(new PdfPCell(new Phrase(val.Value)) { Border = 0 });

                        //if (n > 36)
                        //{
                        //    pdfDoc.NewPage();
                        //    break;
                        //}
                    }

                    //table.AddCell(item.Val);
                    table.AddCell(pdfPTable);
                    if (n % 36 == 0)
                    {
                        pdfDoc.Add(new Paragraph("Hello"));
                        pdfDoc.NewPage();
                        //break;
                    }
                }
            }
            //table.WriteSelectedRows(0, -1, 300, 300, pcb);
            //table.WriteSelectedRows(1, -1, 40, pdfDoc.PageSize.Height - 30, pdfWriter.DirectContent);
            pdfDoc.Add(table);
            pdfWriter.CloseStream = false;
            pdfDoc.Close();
            Response.Buffer = false;
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=Credit-Card-Report.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.Write(pdfDoc);
            Response.End();

            return View();
        }


        [HttpPost]
        [ActionName("Index_Pdf")]
        public ActionResult Index_Pdf()
        {
            CreatePDF();
            return View();
        }

        private void test()
        {
            Document doc = new Document(PageSize.A4.Rotate());

            using (MemoryStream ms = new MemoryStream())
            {
                //PdfWriter writer = PdfWriter.GetInstance(doc, ms);
                //PageEventHelper pageEventHelper = new PageEventHelper();
                //writer.PageEvent = pageEventHelper;
            }
        }

        private void Create_Pdf()
        {

        }

        private void CreatePDF()
        {
            string fileName = string.Empty;
            DateTime fileCreationDatetime = DateTime.Now;
            fileName = string.Format("{0}.pdf", fileCreationDatetime.ToString(@"yyyyMMdd") + "_" + fileCreationDatetime.ToString(@"HHmmss"));
            string pdfPath = Server.MapPath(@"~\PDFs\") + fileName;

            using (FileStream msReport = new FileStream(pdfPath, FileMode.Create))
            {
                //step 1  
                using (Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 140f, 10f))
                {
                    try
                    {
                        // step 2  
                        PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, msReport);
                        pdfWriter.PageEvent = new ITextEvents(Server.MapPath("~/Content/Upload/salma.jpg"));

                        //open the stream   
                        pdfDoc.Open();

                        for (int i = 0; i < 10; i++)
                        {
                            Paragraph para = new Paragraph("Hello world. Checking Header Footer", new Font(Font.FontFamily.HELVETICA, 22));
                            para.Alignment = Element.ALIGN_CENTER;
                            pdfDoc.Add(para);
                            pdfDoc.NewPage();
                        }

                        pdfDoc.Close();
                        Response.Buffer = true;
                        Response.ContentType = "application/pdf";
                        Response.AddHeader("content-disposition", "attachment;filename=Credit-Card-Report.pdf");
                        Response.Cache.SetCacheability(HttpCacheability.NoCache);
                        Response.Write(pdfDoc);
                        Response.End();
                    }
                    catch (Exception ex)
                    {
                        //handle exception  
                    }
                    finally
                    {
                    }
                }
            }
        }



    }

    public class ITextEvents : PdfPageEventHelper
    {
        // This is the contentbyte object of the writer  
        PdfContentByte cb;

        public ITextEvents(string path)
        {
            this.path = path;
        }


        string path { get; set; }

        // we will put the final number of pages in a template  
        PdfTemplate headerTemplate, footerTemplate;

        // this is the BaseFont we are going to use for the header / footer  
        BaseFont bf = null;

        // This keeps track of the creation time  
        DateTime PrintTime = DateTime.Now;

        #region Fields  
        private string _header;
        #endregion

        #region Properties  
        public string Header
        {
            get { return _header; }
            set { _header = value; }
        }
        #endregion

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            try
            {
                PrintTime = DateTime.Now;
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb = writer.DirectContent;
                headerTemplate = cb.CreateTemplate(100, 100);
                footerTemplate = cb.CreateTemplate(50, 50);
            }
            catch (DocumentException de)
            {
            }
            catch (System.IO.IOException ioe)
            {
            }
        }

        public override void OnEndPage(iTextSharp.text.pdf.PdfWriter writer, iTextSharp.text.Document document)
        {
            base.OnEndPage(writer, document);
            iTextSharp.text.Font baseFontNormal = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
            iTextSharp.text.Font baseFontNormal1 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLUE);
            iTextSharp.text.Font baseFontNormal2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLUE);
            iTextSharp.text.Font baseFontNormal3 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.ITALIC, iTextSharp.text.BaseColor.BLUE);
            iTextSharp.text.Font baseFontBig = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
            PdfPCell p1Header = new PdfPCell(new Phrase("Hoang Tam", baseFontNormal1)) { Border = 0 };
            PdfPCell p2Header = new PdfPCell(new Phrase("C14B Street No.9, Le Minh Xuan Industrial Park, Tan Nhut Ward, Binh Chanh District,\nHo Chi Minh City, Viet Nam", baseFontNormal2)) { Border = 0 };
            PdfPCell p3Header = new PdfPCell(new Phrase("Tel: +84 (0) 28 3835 2741", baseFontNormal3)) { Border = 0 };

            ////Table
            //PdfPTable table = new PdfPTable(2);
            //table.WidthPercentage = 100;
            ////0=Left, 1=Centre, 2=Right
            //table.HorizontalAlignment = 0;
            //table.SpacingBefore = 20f;
            //table.SpacingAfter = 30f;

            ////Cell no 1
            //PdfPCell cell = new PdfPCell();
            //cell.Border = 0;
            Image image = Image.GetInstance(path);
            image.ScaleAbsolute(90, 70);
            //cell.AddElement(image);
            //table.AddCell(cell);

            //Create PdfTable object  
            PdfPTable pdfTab = new PdfPTable(2);
            PdfPTable pdfTab1 = new PdfPTable(1);
            pdfTab1.AddCell(p1Header);
            pdfTab1.AddCell(p2Header);
            pdfTab1.AddCell(p3Header);
            //We will have to create separate cells to include image logo and 2 separate strings  
            //Row 1  
            PdfPCell pdfCell1 = new PdfPCell(pdfTab1, new PdfPCell()
            {
                Border = 0,
                PaddingRight = 20
                //BackgroundColor = BaseColor.BLUE
            });
            PdfPCell pdfCell2 = new PdfPCell(image);
            //PdfPCell pdfCell3 = new PdfPCell(image);
            String text = "Page " + writer.PageNumber + " of ";

            //Add paging to header  
            {
                //cb.BeginText();
                //cb.SetFontAndSize(bf, 12);
                //cb.SetTextMatrix(document.PageSize.GetRight(200), document.PageSize.GetTop(45));
                //cb.ShowText(text);
                //cb.EndText();
                //float len = bf.GetWidthPoint(text, 12);
                ////Adds "12" in Page 1 of 12  
                //cb.AddTemplate(headerTemplate, document.PageSize.GetRight(200) + len, document.PageSize.GetTop(45));
            }
            //Add paging to footer  
            {
                cb.BeginText();
                cb.SetFontAndSize(bf, 12);
                cb.SetTextMatrix(document.PageSize.GetRight(180), document.PageSize.GetBottom(30));
                cb.ShowText(text);
                cb.EndText();
                float len = bf.GetWidthPoint(text, 12);
                cb.AddTemplate(footerTemplate, document.PageSize.GetRight(180) + len, document.PageSize.GetBottom(30));
            }

            //Row 2  
            PdfPCell pdfCell4 = new PdfPCell(new Phrase("Sub Header Description", baseFontNormal));

            //Row 3   
            PdfPCell pdfCell5 = new PdfPCell(new Phrase("Date:" + PrintTime.ToShortDateString(), baseFontBig));
            PdfPCell pdfCell6 = new PdfPCell();
            PdfPCell pdfCell7 = new PdfPCell(new Phrase("TIME:" + string.Format("{0:t}", DateTime.Now), baseFontBig));

            //set the alignment of all three cells and set border to 0  
            pdfCell1.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfCell2.HorizontalAlignment = Element.ALIGN_RIGHT;
            //pdfCell3.HorizontalAlignment = Element.ALIGN_RIGHT;
            pdfCell4.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfCell5.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfCell6.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfCell7.HorizontalAlignment = Element.ALIGN_CENTER;

            pdfCell1.VerticalAlignment = Element.ALIGN_TOP;
            pdfCell2.VerticalAlignment = Element.ALIGN_TOP;
            //pdfCell3.VerticalAlignment = Element.ALIGN_TOP;
            pdfCell4.VerticalAlignment = Element.ALIGN_TOP;
            pdfCell5.VerticalAlignment = Element.ALIGN_MIDDLE;
            pdfCell6.VerticalAlignment = Element.ALIGN_MIDDLE;
            pdfCell7.VerticalAlignment = Element.ALIGN_MIDDLE;

            pdfCell4.Colspan = 3;

            pdfCell1.Border = 0;
            pdfCell2.Border = 0;
            //pdfCell3.Border = 0;
            pdfCell4.Border = 0;
            pdfCell5.Border = 0;
            pdfCell6.Border = 0;
            pdfCell7.Border = 0;

            //pdfCell4.BorderWidthLeft = 1f;
            //pdfCell5.BorderWidthLeft = 1f;

            //add all three cells into PdfTable  
            pdfTab.AddCell(pdfCell1);
            pdfTab.AddCell(pdfCell2);
            //pdfTab.AddCell(pdfCell3);
            //pdfTab.AddCell(pdfCell4);
            //pdfTab.AddCell(pdfCell5);
            //pdfTab.AddCell(pdfCell6);
            //pdfTab.AddCell(pdfCell7);

            pdfTab.TotalWidth = document.PageSize.Width - 80f;
            pdfTab.WidthPercentage = 70;
            //pdfTab.HorizontalAlignment = Element.ALIGN_CENTER;      

            //call WriteSelectedRows of PdfTable. This writes rows from PdfWriter in PdfTable  
            //first param is start row. -1 indicates there is no end row and all the rows to be included to write  
            //Third and fourth param is x and y position to start writing  
            pdfTab.WriteSelectedRows(0, -1, 40, document.PageSize.Height - 30, writer.DirectContent);
            //set pdfContent value  

            //Move the pointer and draw line to separate header section from rest of page  
            cb.MoveTo(40, document.PageSize.Height - 120);
            cb.LineTo(document.PageSize.Width - 40, document.PageSize.Height - 120);
            cb.Stroke();

            //Move the pointer and draw line to separate footer section from rest of page  
            cb.MoveTo(40, document.PageSize.GetBottom(50));
            cb.LineTo(document.PageSize.Width - 40, document.PageSize.GetBottom(50));
            cb.Stroke();
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);

            headerTemplate.BeginText();
            headerTemplate.SetFontAndSize(bf, 12);
            headerTemplate.SetTextMatrix(0, 0);
            headerTemplate.ShowText((writer.PageNumber - 1).ToString());
            headerTemplate.EndText();

            footerTemplate.BeginText();
            footerTemplate.SetFontAndSize(bf, 12);
            footerTemplate.SetTextMatrix(0, 0);
            footerTemplate.ShowText((writer.PageNumber - 1).ToString());
            footerTemplate.EndText();
        }
    }
}