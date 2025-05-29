using Newtonsoft.Json;
using PersonelTakip.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PersonelTakip.Controllers
{
    public class BaslangicController : Controller
    {
        // GET: Baslangic
        PersonelDBEntities db = new PersonelDBEntities();
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Harita()
        {
            var personelVeri = db.Tbl_Personel
                .GroupBy(p => p.PerSehirID)
                .Select(g => new
                {
                    SehirId = g.Key,
                    Sayisi = g.Count()
                }).ToList();

            var iller = db.iller.ToList();

            var data = iller.Select(i => new {
                plaka = i.id.ToString("D2"),
                sehirAdi = i.sehir,
                sayi = personelVeri.FirstOrDefault(p => p.SehirId == i.id)?.Sayisi ?? 0
            });
            ViewBag.SehirVeri = JsonConvert.SerializeObject(data);
            ViewBag.ToplamPersonel = db.Tbl_Personel.Count();
            ViewBag.ToplamDrone = db.Tbl_Personel.Select(p => p.PerKursID).Distinct().Count();

            return View();
        }
        public ActionResult IlDetayModal(string plaka)
        {
            // plaka parametresiyle şehir bilgilerini al
            var sehirId = Convert.ToInt32(plaka);
            var personeller = db.Tbl_Personel
                .Where(p => p.PerSehirID == sehirId)
                .ToList();

            return PartialView("_IlDetayModal", personeller);
        }
        public ActionResult IlDetay(string plaka)
        {
            int plakaKodu = int.Parse(plaka);
            var personeller = db.Tbl_Personel
                .Where(p => p.PerSehirID == plakaKodu)
                .ToList();

            return PartialView("_IlDetay", personeller);
        }
    }
}