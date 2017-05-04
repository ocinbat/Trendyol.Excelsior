using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Trendyol.Excelsior;
using Trendyol.Excelsior.Web;
using WebApplication.Models;

namespace WebApplication.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public ActionResult Upload()
        {
            if (Request.Files.Count > 0)
            {
                var file = Request.Files[0];

                if (file != null && file.ContentLength > 0)
                {
                    IExcelsior excelsior = new Excelsior();
                    IEnumerable<Person> persons = excelsior.Listify<Person>(file, true);
                }
            }

            return RedirectToAction("Index");
        }

        [HttpGet]
        public FileResult Download()
        {
            IExcelsior excelsior = new Excelsior();

            List<Person> persons = GetPersons();

            byte[] bytes = excelsior.Excelify(persons, true);
            return File(bytes, "application/vnd.ms-excel", "persons.xlsx");
        }

        private List<Person> GetPersons()
        {
            List<Person> persons = new List<Person>();

            Person person = new Person()
            {
                Id = 1,
                EmploymentStartDate = DateTime.Now.AddDays(-55),
                Name = "Ömer Cinbat",
                RowNumber = 1,
                Decimal = 400.22M,
                Float = 0.18F,
                Double = 0.12D
            };
            persons.Add(person);

            return persons;
        }
    }
}