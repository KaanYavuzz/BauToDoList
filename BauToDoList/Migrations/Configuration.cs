namespace BauToDoList.Migrations
{
    using BauToDoList.Models;
    using Microsoft.AspNet.Identity.EntityFramework;
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;
    using System.Threading.Tasks;
    

    internal sealed class Configuration : DbMigrationsConfiguration<BauToDoList.Models.appDbContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = true;
            AutomaticMigrationDataLossAllowed = true;
        }

        protected override void Seed(BauToDoList.Models.appDbContext context)
        {
            if (!context.Categories.Any())
            {
                //Ýlk kullanýmda DB boþ oluþacaðýn dan sorun yaþamamak için kullanýcý oluþturma iþlemini ve örnek olsun diye de categories tablosuna veri ekleme örneðini bu alan da yapýyoruz. Aþaðýdaki kodlar Categories tablosuna ve ApplicationUser tarafýna ilk kayýtlarýmýzý ekliyoruz.
                context.Categories.Add(new Models.Category() { Name = "Deneme", CreateDate = DateTime.Now, CreatedBy = "username", UpdateDate = DateTime.Now, UpdatedBy = "username" });
                var store = new UserStore<ApplicationUser>(context);
                var manager = new ApplicationUserManager(store);
                var user = new Models.ApplicationUser() { Email = "ahmetdeligoz@gmail.com", UserName = "ahmetdeligoz@gmail.com" };
                Task<Microsoft.AspNet.Identity.IdentityResult> task = Task.Run(() => manager.CreateAsync(user, "ahmet11235"));

                // Will block until the task is completed...
                var result = task.Result;


                context.SaveChanges();
            }
        }
    }
}
