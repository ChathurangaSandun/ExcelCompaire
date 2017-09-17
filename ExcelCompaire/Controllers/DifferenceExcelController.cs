using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ExcelCompaire.Controllers
{
    public class DifferenceExcelController : Controller
    {
        // GET: DifferenceExcel
        public ActionResult Index()
        {

            var planSheet = TempData["planWorksheet"] as IXLWorksheet;
            var productSheet = TempData["productWorksheet"] as IXLWorksheet;








            ViewBag.planSheetName = planSheet.Name;

            return View();
        }

        // GET: DifferenceExcel/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: DifferenceExcel/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: DifferenceExcel/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: DifferenceExcel/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: DifferenceExcel/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: DifferenceExcel/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: DifferenceExcel/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
