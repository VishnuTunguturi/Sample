using ClosedXML.Excel;
using Sample.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Sample.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var model = GetEmployee();
            return View(model);
        }

        /// <summary>
        /// Get Employee List
        /// </summary>
        /// <returns></returns>
        List<EmployeeData> GetEmployee()
        {
            var emp = new List<EmployeeData>()
                {
                new EmployeeData()
                {EmployeeId=1,FirstName="Rakesh",LastName="Kalluri",Email="raki.kalluri@gmail.com",Salary=30000,Company="Summit",Dept="IT"},
                new EmployeeData()
                {EmployeeId=2,FirstName="Naresh",LastName="C",Email="Naresh.C@gmail.com",Salary=50000,Company="IBM",Dept="IT"},
                new EmployeeData()
                {EmployeeId=3,FirstName="Madhu",LastName="K",Email="Madhu.K@gmail.com",Salary=20000,Company="HCl",Dept="IT"},
                new EmployeeData()
                {EmployeeId=4,FirstName="Ali",LastName="MD",Email="Ali.MD@gmail.com",Salary=26700,Company="Tech Mahindra",Dept="BPO"},
                new EmployeeData()
                {EmployeeId=5,FirstName="Chithu",LastName="Raju",Email="Chithu.Raju@gmail.com",Salary=25000,Company="Dell",Dept="BPO"},
                new EmployeeData()
                {EmployeeId=6,FirstName="Nani",LastName="Kumar",Email="Nani.Kumar@gmail.com",Salary=24500,Company="Infosys",Dept="BPO"},

                };
            return emp;
        }

        /// <summary>
        /// Export To Excel
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcel()
        {
            var data = this.GetEmployee();

            DataTable dt = new DataTable();
            // Add Column value
            dt.Columns.Add("EmployeeId");
            dt.Columns.Add("FirstName");
            dt.Columns.Add("LastName");
            dt.Columns.Add("Email");
            dt.Columns.Add("Salary");
            dt.Columns.Add("Company");
            dt.Columns.Add("Dept");

            //  add each of the data rows to the table
            foreach (var row in data)
            {
                DataRow dr;
                dr = dt.NewRow();
                dr[0] = row.EmployeeId;
                dr[1] = row.FirstName;
                dr[2] = row.LastName;
                dr[3] = row.Email;
                dr[4] = row.Salary;
                dr[5] = row.Company;
                dr[6] = row.Dept;
                dt.Rows.Add(dr);
            }


            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Test");
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;

                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename= EmployeeReport.xlsx");

                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
            return RedirectToAction("Index");
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
    }
}