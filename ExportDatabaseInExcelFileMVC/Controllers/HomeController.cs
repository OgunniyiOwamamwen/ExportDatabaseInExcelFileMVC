using System;
using System.Collections.Generic;
using System.Linq;
using ExportDatabaseInExcelFileMVC.Models;
using System.Data.Entity;
using System.Web;
using System.Web.Mvc;
using ExportDatabaseInExcelFileMVC_Entities;
using OfficeOpenXml;
using System.Drawing;

namespace ExportDatabaseInExcelFileMVC.Controllers
{
    public class HomeController : Controller
    {
        // Entities Connetion
        PrixDBEntities db = new PrixDBEntities();

        public ActionResult Index()
        {
            List<EmployeeViewModel> employee = db.EmployeeInfoes.Select(x=> new EmployeeViewModel
            {
                EmployeeId = x.EmployeeId,
                EmployeeName = x.EmployeeName,
                Email = x.Email,
                Phone = x.Phone,
                Experience = x.Experience
            }).ToList();

            return View(employee);
        }
        public void ExportToExcel()
        {
            List<EmployeeViewModel> employee = db.EmployeeInfoes.Select(x => new EmployeeViewModel
            {
                EmployeeId = x.EmployeeId,
                EmployeeName = x.EmployeeName,
                Email = x.Email,
                Phone = x.Phone,
                Experience = x.Experience
            }).ToList();
            ExcelPackage pck = new ExcelPackage();          
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");
            // if error change' ExcelWorksheet ws' exmple var ws
            ws.Cells["A1"].Value = "PRIX";
            ws.Cells["B1"].Value = "EMPLOYEE DATA";

            ws.Cells["A2"].Value = "Report";
            ws.Cells["B2"].Value = "All Data";

            ws.Cells["A3"].Value = "Date";
            ws.Cells["B3"].Value = string.Format("{0:dd mmmm yyyy} at {0:H: mm tt}",DateTimeOffset.Now);

            ws.Cells["A6"].Value = "EmployeeId";
            ws.Cells["B6"].Value = "EmployeeName";
            ws.Cells["C6"].Value = "Email";
            ws.Cells["D6"].Value = "Phone";
            ws.Cells["E6"].Value = "Experience";

            int rowStar = 7;
            foreach(var item in employee)
            {
                if (item.Experience < 5)
                {
                    ws.Row(rowStar).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Row(rowStar).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("pink")));
                }
                ws.Cells[string.Format("A{0}", rowStar)].Value = item.EmployeeId;
                ws.Cells[string.Format("B{0}", rowStar)].Value = item.EmployeeName;
                ws.Cells[string.Format("C{0}", rowStar)].Value = item.Email;
                ws.Cells[string.Format("D{0}", rowStar)].Value = item.Phone;
                ws.Cells[string.Format("E{0}", rowStar)].Value = item.Experience;
                rowStar++;
            }
            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename=" + "ExcelReport.xlsx");
            Response.BinaryWrite(pck.GetAsByteArray());
            Response.End();
        }
    }
}