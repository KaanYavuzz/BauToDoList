using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using BauToDoList.Models;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;

namespace BauToDoList.Controllers
{
    public class ToDoItemsController : Controller
    {
        private appDbContext db = new appDbContext();

        // GET: ToDoItems
        public async Task<ActionResult> Index()
        {
            var toDoItems = db.ToDoItems.Include(t => t.Category).Include(t => t.Customer).Include(t => t.Department).Include(t => t.Manager).Include(t => t.Organizator).Include(t => t.Side);
            return View(await toDoItems.ToListAsync());
        }

        // GET: ToDoItems/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ToDoItem toDoItem = await db.ToDoItems.FindAsync(id);
            if (toDoItem == null)
            {
                return HttpNotFound();
            }
            return View(toDoItem);
        }

        // GET: ToDoItems/Create
        public ActionResult Create()
        {
            var toDoItem = new ToDoItem();
            toDoItem.MeetingDate = DateTime.Now;
            toDoItem.MeetingHour = DateTime.Now;
            toDoItem.FinishDate = DateTime.Now;
            toDoItem.PlannedDate = DateTime.Now;
            toDoItem.PlannedHour = DateTime.Now;
            toDoItem.ReviseDate = DateTime.Now;
            toDoItem.ReviseHour = DateTime.Now;
            toDoItem.ScheduledOrganizationDate = DateTime.Now;
            toDoItem.ScheduledOrganizationHour = DateTime.Now;
            ViewBag.CategoryId = new SelectList(db.Categories, "Id", "Name");
            ViewBag.CustomerId = new SelectList(db.Customers, "Id", "Name");
            ViewBag.DepartmentId = new SelectList(db.Departments, "Id", "Name");
            ViewBag.ManagerId = new SelectList(db.Contacts, "Id", "FirstName");
            ViewBag.OrganizatorId = new SelectList(db.Contacts, "Id", "FirstName");
            ViewBag.SideId = new SelectList(db.Sides, "Id", "Name");

            return View(toDoItem);
        }

        // POST: ToDoItems/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "Id,Title,Description,Status,CategoryId,Attachment,DepartmentId,SideId,CustomerId,ManagerId,OrganizatorId,MeetingDate,MeetingHour,PlannedDate,PlannedHour,FinishDate,FinishHour,ReviseDate,ReviseHour,ConversationSubject,SupporterCompany,SupporterDoctor,ConversationAttendeeCount,ScheduledOrganizationDate,ScheduledOrganizationHour,MailingSubjects,PosterSubject,PosterCount,Elearning,TypesOfScans,AsoCountInScans,TypesOfOrganization,AsoCountInOrganization,TypesOfVaccinationOrganization,AsoCountInVaccinationOrganization,AmountOfCompensantionForPoster,CorporateProductivityReport,CreatedBy,CreateDate,UpdatedBy,UpdateDate")] ToDoItem toDoItem)
        {
            if (ModelState.IsValid)
            {
                toDoItem.CreateDate = DateTime.Now;
                toDoItem.CreatedBy = User.Identity.Name;
                toDoItem.UpdateDate = DateTime.Now;
                toDoItem.UpdatedBy = User.Identity.Name;
                db.ToDoItems.Add(toDoItem);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            ViewBag.CategoryId = new SelectList(db.Categories, "Id", "Name", toDoItem.CategoryId);
            ViewBag.CustomerId = new SelectList(db.Customers, "Id", "Name", toDoItem.CustomerId);
            ViewBag.DepartmentId = new SelectList(db.Departments, "Id", "Name", toDoItem.DepartmentId);
            ViewBag.ManagerId = new SelectList(db.Contacts, "Id", "FirstName", toDoItem.ManagerId);
            ViewBag.OrganizatorId = new SelectList(db.Contacts, "Id", "FirstName", toDoItem.OrganizatorId);
            ViewBag.SideId = new SelectList(db.Sides, "Id", "Name", toDoItem.SideId);
            return View(toDoItem);
        }

        // GET: ToDoItems/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ToDoItem toDoItem = await db.ToDoItems.FindAsync(id);
            if (toDoItem == null)
            {
                return HttpNotFound();
            }
            ViewBag.CategoryId = new SelectList(db.Categories, "Id", "Name", toDoItem.CategoryId);
            ViewBag.CustomerId = new SelectList(db.Customers, "Id", "Name", toDoItem.CustomerId);
            ViewBag.DepartmentId = new SelectList(db.Departments, "Id", "Name", toDoItem.DepartmentId);
            ViewBag.ManagerId = new SelectList(db.Contacts, "Id", "FirstName", toDoItem.ManagerId);
            ViewBag.OrganizatorId = new SelectList(db.Contacts, "Id", "FirstName", toDoItem.OrganizatorId);
            ViewBag.SideId = new SelectList(db.Sides, "Id", "Name", toDoItem.SideId);
            return View(toDoItem);
        }

        // POST: ToDoItems/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "Id,Title,Description,Status,CategoryId,Attachment,DepartmentId,SideId,CustomerId,ManagerId,OrganizatorId,MeetingDate,MeetingHour,PlannedDate,PlannedHour,FinishDate,FinishHour,ReviseDate,ReviseHour,ConversationSubject,SupporterCompany,SupporterDoctor,ConversationAttendeeCount,ScheduledOrganizationDate,ScheduledOrganizationHour,MailingSubjects,PosterSubject,PosterCount,Elearning,TypesOfScans,AsoCountInScans,TypesOfOrganization,AsoCountInOrganization,TypesOfVaccinationOrganization,AsoCountInVaccinationOrganization,AmountOfCompensantionForPoster,CorporateProductivityReport,CreatedBy,CreateDate,UpdatedBy,UpdateDate")] ToDoItem toDoItem)
        {
            if (ModelState.IsValid)
            {
                toDoItem.UpdateDate = DateTime.Now;
                toDoItem.UpdatedBy = User.Identity.Name;
                db.Entry(toDoItem).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            ViewBag.CategoryId = new SelectList(db.Categories, "Id", "Name", toDoItem.CategoryId);
            ViewBag.CustomerId = new SelectList(db.Customers, "Id", "Name", toDoItem.CustomerId);
            ViewBag.DepartmentId = new SelectList(db.Departments, "Id", "Name", toDoItem.DepartmentId);
            ViewBag.ManagerId = new SelectList(db.Contacts, "Id", "FirstName", toDoItem.ManagerId);
            ViewBag.OrganizatorId = new SelectList(db.Contacts, "Id", "FirstName", toDoItem.OrganizatorId);
            ViewBag.SideId = new SelectList(db.Sides, "Id", "Name", toDoItem.SideId);
            return View(toDoItem);
        }

        // GET: ToDoItems/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ToDoItem toDoItem = await db.ToDoItems.FindAsync(id);
            if (toDoItem == null)
            {
                return HttpNotFound();
            }
            return View(toDoItem);
        }

        // POST: ToDoItems/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            ToDoItem toDoItem = await db.ToDoItems.FindAsync(id);
            db.ToDoItems.Remove(toDoItem);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
        public void ExportToExcel()
        {
            var grid = new GridView();
            grid.DataSource = from data in db.ToDoItems.ToList()
                              select new
                              {
                                  Baslik = data.Title,
                                  Aciklama = data.Description,
                                  Durum = data.Status,
                                  Kategori = data.CategoryId,
                                  DosyaEki = data.Attachment,
                                  Departman = data.DepartmentId,
                                  Taraf = data.SideId,
                                  Musteri = data.Customer,
                                  Yonetici = data.ManagerId,
                                  Organizator = data.OrganizatorId,
                                  ToplantiTarihi = data.MeetingDate,
                                  ToplantiSaati = data.MeetingHour,
                                  PlanlananTarih = data.PlannedDate,
                                  PlanlananSaat = data.PlannedHour,
                                  BitisTarihi = data.FinishDate,
                                  BitisSaati = data.FinishHour,
                                  RevizeTarihi = data.ReviseDate,
                                  RevizeSaati = data.ReviseHour,
                                  GorusmeKonusu = data.ConversationSubject,
                                  SponsorFirma = data.SupporterCompany,
                                  DestekleyenDoktor = data.SupporterDoctor,
                                  GorusmeyeKatilimciSayisi = data.ConversationAttendeeCount,
                                  PlanlanmisOrganizasyonTarihi = data.ScheduledOrganizationDate,
                                  PlanlanmisOrganizasyonSaati = data.ScheduledOrganizationHour,
                                  MailKonulari = data.MailingSubjects,
                                  PosterKonusu = data.PosterSubject,
                                  UzaktanEgitim = data.Elearning,
                                  TaramaTurleri = data.TypesOfScans,
                                  TaranmisAsoSayisi = data.AsoCountInScans,
                                  OrganizasyonTurleri = data.TypesOfOrganization,
                                  OrganizasyonlardakiAsoSayisi = data.AsoCountInOrganization,
                                  AsiOrganizasyonTurleri = data.TypesOfVaccinationOrganization,
                                  AsiOrganizasyolarindakiAsoSayisi = data.AsoCountInVaccinationOrganization,
                                  AfisIcinButceMiktari = data.AmountOfCompensantionForPoster,
                                  KurumsalVerimlilikRaporu = data.CorporateProductivityReport,
                                  OlusturmaTarihi = data.CreateDate,
                                  OlusturanKullanici = data.CreatedBy,
                                  GuncellemeTarihi = data.UpdateDate,
                                  GuncelleyenKullanıcı = data.UpdatedBy,
                              };

            grid.DataBind();
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename=Yapılacaklar.xls");
            Response.ContentType = "application/ms-excel";
            Response.ContentEncoding = System.Text.Encoding.Unicode;
            Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);
            grid.RenderControl(hw);
            Response.Write(sw.ToString());
            Response.End();
        }

        public void ExportToCsv()
        {
            StringWriter sw = new StringWriter();
            sw.WriteLine("Title,Description,Status,CategoryId,Attachment,DepartmentId,SideId,CustomerId,ManagerId,OrganizatorId,MeetingDate,MeetingHour,PlannedDate,PlannedHour,FinishDate,FinishHour,ReviseDate,ReviseHour,ConversationSubject,SupporterCompany,SupporterDoctor,ConversationAttendeeCount,ScheduledOrganizationDate,ScheduledOrganizationHour,MailingSubjects,PosterSubject,PosterCount,Elearning,TypesOfScans,AsoCountInScans,TypesOfOrganization,AsoCountInOrganization,TypesOfVaccinationOrganization,AsoCountInVaccinationOrganization,AmountOfCompensantionForPoster,CorporateProductivityReport,CreatedBy,CreateDate,UpdatedBy,UpdateDate");
            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment;filename=Yapılacaklar.csv");
            Response.ContentType = "text/csv";
            Response.ContentEncoding = System.Text.Encoding.Unicode;
            Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());
            var todoitem = db.ToDoItems;
            foreach (var todoitems in todoitem)
            {
                sw.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28},{29},{30},{31},{32},{33},{34},{35},{36},{37},{38}",
                    todoitems.Title,
                    todoitems.Description,
                    todoitems.Status,
                    todoitems.CategoryId,
                    todoitems.Attachment,
                    todoitems.DepartmentId,
                    todoitems.SideId,
                    todoitems.CustomerId,
                    todoitems.ManagerId,
                    todoitems.OrganizatorId,
                    todoitems.MeetingDate,
                    todoitems.MeetingHour,
                    todoitems.PlannedDate,
                    todoitems.PlannedHour,
                    todoitems.FinishDate,
                    todoitems.FinishHour,
                    todoitems.ReviseDate,
                    todoitems.ReviseHour,
                    todoitems.ConversationSubject,
                    todoitems.SupporterCompany,
                    todoitems.SupporterDoctor,
                    todoitems.ConversationAttendeeCount,
                    todoitems.ScheduledOrganizationDate,
                    todoitems.ScheduledOrganizationHour,
                    todoitems.MailingSubjects,
                    todoitems.PosterSubject,
                    todoitems.Elearning,
                    todoitems.TypesOfScans,
                    todoitems.AsoCountInScans,
                    todoitems.TypesOfOrganization,
                    todoitems.AsoCountInOrganization,
                    todoitems.TypesOfVaccinationOrganization,
                    todoitems.AsoCountInVaccinationOrganization,
                    todoitems.AmountOfCompensantionForPoster,
                    todoitems.CorporateProductivityReport,
                    todoitems.CreateDate,
                    todoitems.CreatedBy,
                    todoitems.UpdateDate,
                    todoitems.UpdatedBy
                    ));
               
            }
            Response.Write(sw.ToString());
            Response.End();
        }
    }
}
