using PersonelTakip.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PersonelTakip.Controllers
{
    public class KullaniciController : Controller
    {
        // GET: Kullanici

        PersonelDBEntities db = new PersonelDBEntities();
        [HttpGet]
        public ActionResult Giris()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Giris(Tbl_Admin admin)
        {
            var kullanici = db.Tbl_Admin
                              .FirstOrDefault(x => x.AdminSicil == admin.AdminSicil && x.AdminSifre == admin.AdminSifre);

            if (kullanici != null)
            {
                Session["kullanici"] = kullanici.AdminSicil;
                return RedirectToAction("Liste", "Personel");
            }
            else
            {
                ViewBag.Hata = "Kullanıcı adı veya şifre yanlış!";
                return View();
            }
        }

        public ActionResult Cikis()
        {
            Session.Abandon();
            return RedirectToAction("Giris");
        }
    }
}
