using iTextSharp.text.pdf;
using iTextSharp.text;
using OfficeOpenXml;
using PersonelTakip.Models;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;


namespace PersonelTakip.Controllers
{
    public class RaporController : Controller
    {
        // GET: Rapor
        PersonelDBEntities db = new PersonelDBEntities();
        public ActionResult Index()
        {
            ViewBag.Sehirler = new SelectList(db.iller, "sehirID", "sehir");
            ViewBag.Kurslar = new SelectList(db.Tbl_PersonelKurs, "ID", "PersonelKursAd");
            ViewBag.Rutbeler = db.Tbl_Personel.Select(x => x.PerRutbe).Distinct().ToList();
            return View();
        }
        [HttpPost]
        public ActionResult Sonuc(int? sehirID, int? kursID, string rutbe)
        {
            var sorgu = db.Tbl_Personel.AsQueryable();

            if (sehirID != null)
                sorgu = sorgu.Where(p => p.PerSehirID == sehirID);

            if (kursID != null)
                sorgu = sorgu.Where(p => p.PerKursID == kursID);

            if (!string.IsNullOrEmpty(rutbe))
                sorgu = sorgu.Where(p => p.PerRutbe == rutbe);

            return View(sorgu.ToList());
        }
        
        public ActionResult PdfRapor()
        {
            var personeller = db.Tbl_Personel.ToList();

            using (MemoryStream stream = new MemoryStream())
            {
                // PDF nesnesi oluştur
                Document pdfDoc = new Document(PageSize.A4, 25, 25, 30, 30);
                PdfWriter writer = PdfWriter.GetInstance(pdfDoc, stream);
                writer.CloseStream = false;

                // Belgeyi aç
                pdfDoc.Open();
                pdfDoc.Add(new Paragraph("Personel Raporu"));
                pdfDoc.Add(new Paragraph(" "));

                // Tablo oluştur (4 sütunlu)
                PdfPTable table = new PdfPTable(4);

                // Başlık hücreleri
                table.AddCell("Ad");
                table.AddCell("Soyad");
                table.AddCell("Sicil");
                table.AddCell("Rütbe");

                // Verileri doldur
                foreach (var p in personeller)
                {
                    table.AddCell(p.PerAd ?? "");
                    table.AddCell(p.PerSoyad ?? "");
                    table.AddCell(p.PerSicil ?? "");
                    table.AddCell(p.PerRutbe ?? "");
                }

                pdfDoc.Add(table);
                pdfDoc.Close();

                byte[] bytes = stream.ToArray();
                return File(bytes, "application/pdf", "PersonelRapor.pdf");
            }
        }
        public ActionResult ExcelRapor()
        {
            var personeller = db.Tbl_Personel.ToList();

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("PersonelListesi");

                // Başlıklar
                sheet.Cells[1, 1].Value = "Ad";
                sheet.Cells[1, 2].Value = "Soyad";
                sheet.Cells[1, 3].Value = "Sicil";
                sheet.Cells[1, 4].Value = "Rütbe";

                int row = 2;
                foreach (var p in personeller)
                {
                    sheet.Cells[row, 1].Value = p.PerAd;
                    sheet.Cells[row, 2].Value = p.PerSoyad;
                    sheet.Cells[row, 3].Value = p.PerSicil;
                    sheet.Cells[row, 4].Value = p.PerRutbe;
                    row++;
                }

                var stream = new MemoryStream(package.GetAsByteArray());
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PersonelRapor.xlsx");
            }
        }
    }
}
            
            
        
    
    

