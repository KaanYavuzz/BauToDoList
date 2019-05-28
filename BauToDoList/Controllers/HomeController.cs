using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BauToDoList.Models;

namespace BauToDoList.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        appDbContext db = new appDbContext();
        public ActionResult Index()
        {
            ViewBag.CustomerCount = db.Customers.Count();
            ViewBag.StatusNewCount = db.ToDoItems.Count(t=>t.Status==Status.New);
            ViewBag.StatusWaitingCount = db.ToDoItems.Count(t => t.Status == Status.Waiting);
            ViewBag.StatusCompletedCount = db.ToDoItems.Count(t => t.Status == Status.Complated);

            return View();
        }

        public ActionResult Kurumsal()
        {
            var department = new Department();
            department.Name = "Pazarlama";
            return View(department);
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";
            var contact = new Contact();
            return View(contact);
        }
        [HttpPost]
        public ActionResult Contact(Contact contact)
        {
            if (ModelState.IsValid)
            {
                using (var db = new appDbContext())
                {
                    contact.CreateDate = DateTime.Now;
                    contact.CreatedBy = User.Identity.Name;
                    contact.UpdateDate = DateTime.Now;
                    contact.UpdatedBy = User.Identity.Name;

                    db.Contacts.Add(contact);
                    db.SaveChanges();
                    Session["StatusMessage"] = "Kişi formu başarıyla kaydedildi";
                    return RedirectToAction("Index");
                }
            }
            return View(contact);
        }
    }
}