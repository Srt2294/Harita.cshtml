using iTextSharp.text.pdf;
using iTextSharp.text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PersonelTakip.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;




namespace PersonelTakip.Controllers
{
    public class PersonelController : Controller
    {
        // GET: Personel
        PersonelDBEntities db = new PersonelDBEntities();
        /* public ActionResult Liste()
         {
             if (Session["kullanici"] == null)
                 return RedirectToAction("Giris", "kullanici");
             var personel = db.Tbl_Personel.ToList();
             return View(personel);

         }*/
        // GET: Personel/Ekle
        public ActionResult Ekle()
        {
            ViewBag.PerSehirID = new SelectList(db.iller.ToList(), "id", "sehir");
            ViewBag.PerKursID = new SelectList(db.Tbl_PersonelKurs.ToList(), "PersonelKursID", "PersonelKursAd");
            return View();
        }

        // POST: Personel/Ekle
        [HttpPost]
        public ActionResult Ekle(Tbl_Personel personel)
        {

            if (ModelState.IsValid)
            {
                db.Tbl_Personel.Add(personel);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(personel);
        }
        public ActionResult Sil(int id)
        {
            ViewBag.PerSehirID = new SelectList(db.iller.ToList(), "id", "sehir");
            ViewBag.PerKursID = new SelectList(db.Tbl_PersonelKurs.ToList(), "PersonelKursID", "PersonelKursAd");
            
            var personel = db.Tbl_Personel.Find(id);
            if (personel != null)
            {
                db.Tbl_Personel.Remove(personel);
                db.SaveChanges();
            }
            return RedirectToAction("Index");
        }
        // GET: Personel/Guncelle/5
        public ActionResult Guncelle(int id)
        {
            ViewBag.PerSehirID = new SelectList(db.iller.ToList(), "id", "sehir");
            ViewBag.PerKursID = new SelectList(db.Tbl_PersonelKurs.ToList(), "PersonelKursID", "PersonelKursAd");
            
            var personel = db.Tbl_Personel.Find(id);
            return View(personel);
        }

        // POST: Personel/Guncelle/5
        [HttpPost]
        public ActionResult Guncelle(Tbl_Personel personel)
        {
            if (ModelState.IsValid)
            {
                db.Entry(personel).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(personel);
        }
        public ActionResult ExcelRapor()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Personeller");

                sheet.Cells[1, 1].Value = "Ad";
                sheet.Cells[1, 2].Value = "Soyad";
                sheet.Cells[1, 3].Value = "Sicil";
                sheet.Cells[1, 4].Value = "Rütbe";
                sheet.Cells[1, 5].Value = "Şehir";

                int row = 2;
                var personeller = db.Tbl_Personel.Include("iller").ToList();

                foreach (var p in personeller)
                {
                    sheet.Cells[row, 1].Value = p.PerAd;
                    sheet.Cells[row, 2].Value = p.PerSoyad;
                    sheet.Cells[row, 3].Value = p.PerSicil;
                    sheet.Cells[row, 4].Value = p.PerRutbe;
                    sheet.Cells[row, 5].Value = p.iller?.sehir ?? ""; // Şehir adı ekleniyor
                    row++;
                }

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PersonelRapor.xlsx");
            }
        }
        public ActionResult PdfRapor()
        {
            // Personel listesini al (veritabanından vs.)
            var personeller = db.Tbl_Personel.Include("iller").ToList();

            using (MemoryStream stream = new MemoryStream())
            {
                Document pdfDoc = new Document(PageSize.A4, 25, 25, 30, 30);
                PdfWriter.GetInstance(pdfDoc, stream).CloseStream = false;

                pdfDoc.Open();
                pdfDoc.Add(new Paragraph("Personel Raporu"));
                pdfDoc.Add(new Paragraph(" "));

                PdfPTable table = new PdfPTable(5); // Kolon sayısı
                table.AddCell("Ad");
                table.AddCell("Soyad");
                table.AddCell("Sicil");
                table.AddCell("Rütbe");
                table.AddCell("Şehir"); // Yeni ek

                foreach (var p in personeller)
                {
                    table.AddCell(p.PerAd);
                    table.AddCell(p.PerSoyad);
                    table.AddCell(p.PerSicil);
                    table.AddCell(p.PerRutbe);
                    table.AddCell(p.iller?.sehir ?? ""); // Şehir adı
                }

                pdfDoc.Add(table);
                pdfDoc.Close();

                byte[] bytes = stream.ToArray();
                return File(bytes, "application/pdf", "PersonelRapor.pdf");
            }

        }
        public ActionResult Index(int? sehirId, string rutbe)
        {
            var personeller = db.Tbl_Personel
                .Include(p => p.iller)
                .Include(p => p.Tbl_IhaTuru)
                .Include(p => p.Tbl_PersonelKurs)
                .AsQueryable();

            if (sehirId.HasValue)
            {
                personeller = personeller.Where(p => p.PerSehirID == sehirId.Value);
            }

            if (!string.IsNullOrEmpty(rutbe))
            {
                personeller = personeller.Where(p => p.PerRutbe.Contains(rutbe));
            }

            ViewBag.SehirId = new SelectList(db.iller.ToList(), "id", "sehir");
            ViewBag.Rutbe = rutbe;

            return View(personeller.ToList());
        }
    }
}
