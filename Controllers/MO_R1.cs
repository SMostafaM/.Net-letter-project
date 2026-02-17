using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MongoDB.Driver;
using MongoDB.Bson;
using ClosedXML.Excel;
using daftar.Models;

namespace daftar.Controllers
{
    public class MO_R1 : Controller
    {
        private Context context;
        public MO_R1()
        {
            context = new Context();
        }

        public int Chek_cooke()
        {
            int C = 0;
            if (Request.Cookies["token"] != null)
            {
                string cookieValue = Request.Cookies["token"];
                string token = "";
                for (int i = 0; i < 120; i++)
                {
                    token = token + cookieValue[i];
                }
                List<Models.Active_session> acs = new List<Active_session>();
                acs = context.Active_session.Find(a => a.user_token == token).ToList();
                if (acs != null)
                {
                    if (acs[0].active_sesstion == cookieValue)
                    {
                        var cook = new Microsoft.AspNetCore.Http.CookieOptions() { Path = "/", HttpOnly = false, IsEssential = true, Expires = DateTime.Now.AddHours(10) };
                        Response.Cookies.Append("token", cookieValue, cook);
                        C = 1;
                    }
                }

            }

            return C;
        }

        public string chek_Roel()
        {
            security.Coding coding = new security.Coding();
            string R = "";
            string Role = "";
            string cookieValue = Request.Cookies["token"];
            for (int i = 80; i < 120; i++)
            {
                Role = Role + cookieValue[i];
            }

            string rolesha1 = "";

            //List<string> send_list = new List<string>();
            //send_list.Add("0");
            //send_list.Add("0");
            string send_str = "0 0";

            for (int i = 0; i <= 1000; i++)
            {
                rolesha1 = coding.Sha1Sum(i.ToString());
                if (rolesha1 == Role)
                {
                    //send_list.Clear();
                    //send_list.Add("1");
                    //send_list.Add(i.ToString());
                    send_str = "1 " + i.ToString();
                    break;
                }
            }
            return send_str;
            //return send_list;
    
        }

        public IActionResult write_xls(List<Models.Letter> list_to_write)
        {

            DateTime localDate = DateTime.Now;
            string time = localDate.Year.ToString() + "-" + localDate.Month.ToString() + "-" + localDate.Day.ToString();

            //try
            //{
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(time);
                int currentRow = 1;

                worksheet.Cell(currentRow, 1).Value = "سریال نامه";
                worksheet.Cell(currentRow, 2).Value = "عنوان نامه";
                worksheet.Cell(currentRow, 3).Value = "شماره نامه وارده";
                worksheet.Cell(currentRow, 4).Value = "تاریخ ایجاد نامه";
                worksheet.Cell(currentRow, 5).Value = "تاریخ ارجاع به معاونت مربوطه";
                worksheet.Cell(currentRow, 6).Value = "معاونت";
                worksheet.Cell(currentRow, 7).Value = "مهلت پاسخ";
                worksheet.Cell(currentRow, 8).Value = "دفعات پیگیری";
                worksheet.Cell(currentRow, 9).Value = "مهلت های داده شده";
                worksheet.Cell(currentRow, 10).Value = "تاریخ پیگیری های انجام شده";
                worksheet.Cell(currentRow, 11).Value = "وضعیت نامه";
                worksheet.Cell(currentRow, 12).Value = "سریال پاسخ نامه";
                worksheet.Cell(currentRow, 13).Value = "تاریخ پاسخ نامه";


                string last_d = "";
                foreach (var item in list_to_write)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = item.serial;
                    worksheet.Cell(currentRow, 2).Value = item.title;
                    worksheet.Cell(currentRow, 3).Value = item.letter_number;
                    worksheet.Cell(currentRow, 4).Value = item.date_create;
                    worksheet.Cell(currentRow, 5).Value = item.date_send;
                    worksheet.Cell(currentRow, 6).Value = item.place;
                    worksheet.Cell(currentRow, 7).Value = item.day_reserve;
                    worksheet.Cell(currentRow, 8).Value = item.count_followed;
                    if (item.last_day_reserve != null)
                    {
                        foreach (string i in item.last_day_reserve)
                        {
                            last_d = last_d + i.ToString();
                        }
                    }
                    worksheet.Cell(currentRow, 9).Value = last_d;
                    last_d = "";
                    if (item.list_date_fllowed != null)
                    {
                        foreach (string i in item.list_date_fllowed)
                        {
                            last_d = last_d + i.ToString();
                        }
                    }
                    worksheet.Cell(currentRow, 10).Value = last_d;
                    worksheet.Cell(currentRow, 11).Value = item.letter_status;
                    worksheet.Cell(currentRow, 12).Value = item.awnser_serial;
                    worksheet.Cell(currentRow, 13).Value = item.date_awnser;

                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    time = time + ".xls";
                    return File(content, "text/plain", time);
                }
            }
            //}
            //catch (Exception e)
            //{
            //    return BadRequest(e.Message);
            //}


        }


        public IActionResult write_xls_sanction(List<Models.Sanction> list_to_write)
        {

            DateTime localDate = DateTime.Now;
            string time = localDate.Year.ToString() + "-" + localDate.Month.ToString() + "-" + localDate.Day.ToString();

            //try
            //{
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(time);
                int currentRow = 1;

                worksheet.Cell(currentRow, 1).Value = "سریال نامه";
                worksheet.Cell(currentRow, 2).Value = "عنوان جلسه";
                worksheet.Cell(currentRow, 3).Value = "شماره جلسه";
                worksheet.Cell(currentRow, 4).Value = "تاریخ جلسه";
                worksheet.Cell(currentRow, 5).Value = "تاریخ ارجاع به معاونت مربوطه";
                worksheet.Cell(currentRow, 6).Value = "معاونت";
                worksheet.Cell(currentRow, 7).Value = "مهلت پاسخ";
                worksheet.Cell(currentRow, 8).Value = "دفعات پیگیری";
                worksheet.Cell(currentRow, 9).Value = "مهلت های داده شده";
                worksheet.Cell(currentRow, 10).Value = "تاریخ پیگیری های انجام شده";
                worksheet.Cell(currentRow, 11).Value = "وضعیت نامه";
                worksheet.Cell(currentRow, 12).Value = "سریال پاسخ نامه";
                worksheet.Cell(currentRow, 13).Value = "تاریخ پاسخ نامه";
                worksheet.Cell(currentRow, 13).Value = "تاریخ به بارزسی";


                string last_d = "";
                foreach (var item in list_to_write)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = item.serial;
                    worksheet.Cell(currentRow, 2).Value = item.title;
                    worksheet.Cell(currentRow, 3).Value = item.meetingـnumber;
                    worksheet.Cell(currentRow, 4).Value = item.date_create;
                    worksheet.Cell(currentRow, 5).Value = item.date_send;
                    worksheet.Cell(currentRow, 6).Value = item.place;
                    worksheet.Cell(currentRow, 7).Value = item.day_reserve;
                    worksheet.Cell(currentRow, 8).Value = item.count_followed;
                    if (item.last_day_reserve != null)
                    {
                        foreach (string i in item.last_day_reserve)
                        {
                            last_d = last_d + i.ToString();
                        }
                    }
                    worksheet.Cell(currentRow, 9).Value = last_d;
                    last_d = "";
                    if (item.list_date_fllowed != null)
                    {
                        foreach (string i in item.list_date_fllowed)
                        {
                            last_d = last_d + i.ToString();
                        }
                    }
                    worksheet.Cell(currentRow, 10).Value = last_d;
                    worksheet.Cell(currentRow, 11).Value = item.letter_status;
                    worksheet.Cell(currentRow, 12).Value = item.awnser_serial;
                    worksheet.Cell(currentRow, 13).Value = item.date_awnser;
                    worksheet.Cell(currentRow, 13).Value = item.date_send_to_H_B;

                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    time = time + ".xls";
                    return File(content, "text/plain", time);
                }
            }
            //}
            //catch (Exception e)
            //{
            //    return BadRequest(e.Message);
            //}


        }

        // GET: MO

        // GET: Letter
        public IActionResult Index()
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro= chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }

            //solar date
            System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            DateTime dt = DateTime.Now;
            int y = p.GetYear(dt);
            int m = p.GetMonth(dt);
            int d = p.GetDayOfMonth(dt);
            string date_today = y.ToString() + "/" + m.ToString() + "/" + d.ToString();
            ViewBag.time = date_today;

            List<Models.Letter> list_letter = context.Letter.Find(x => x.letter_status == "اتمام مهلت" && x.place == R[1] || x.letter_status == "فوری" && x.place == R[1]).ToList();
            return View(list_letter);
        }
        public IActionResult Index_send_H_B()
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }

            //solar date
            //System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            //DateTime dt = DateTime.Now;
            //int y = p.GetYear(dt);
            //int m = p.GetMonth(dt);
            //int d = p.GetDayOfMonth(dt);
            //string date_today = d.ToString() + "/" + m.ToString() + "/" + y.ToString();
            //ViewBag.time = date_today;


            List<Models.Letter> list_letter = context.Letter.Find(x => x.letter_status == "ارسال به بازرسی" && x.place == R[1]).ToList();
            return View(list_letter);
        }
        public IActionResult Index_in_progres()
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }

            //solar date
            System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            DateTime dt = DateTime.Now;
            int y = p.GetYear(dt);
            int m = p.GetMonth(dt);
            int d = p.GetDayOfMonth(dt);
            string date_today = y.ToString() + "/" + m.ToString() + "/" + d.ToString();
            ViewBag.time = date_today;

            List<Models.Letter> list_letter = context.Letter.Find(x => x.letter_status == "در حال پیگیری" && x.place == R[1] || x.letter_status == "بررسی پاسخ" && x.place == R[1]).ToList();
            return View(list_letter);
        }

        public IActionResult Index_sanction()
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            //solar date
            System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            DateTime dt = DateTime.Now;
            int y = p.GetYear(dt);
            int m = p.GetMonth(dt);
            int d = p.GetDayOfMonth(dt);
            string date_today = y.ToString() + "/" + m.ToString() + "/" + d.ToString();
            ViewBag.time = date_today;

            List<Models.Sanction> list_letter = context.sanction.Find(x => x.letter_status == "اتمام مهلت" && x.place == R[1] || x.letter_status == "فوری" && x.place == R[1]).ToList();
            return View(list_letter);
        }
        public IActionResult Index_send_H_B_sanction()
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }

            //solar date
            //System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            //DateTime dt = DateTime.Now;
            //int y = p.GetYear(dt);
            //int m = p.GetMonth(dt);
            //int d = p.GetDayOfMonth(dt);
            //string date_today = d.ToString() + "/" + m.ToString() + "/" + y.ToString();
            //ViewBag.time = date_today;


            List<Models.Sanction> list_letter = context.sanction.Find(x => x.letter_status == "ارسال به بازرسی" && x.place == R[1]).ToList();
            return View(list_letter);
        }
        public IActionResult Index_in_progres_sanction()
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }

            //solar date
            System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            DateTime dt = DateTime.Now;
            int y = p.GetYear(dt);
            int m = p.GetMonth(dt);
            int d = p.GetDayOfMonth(dt);
            string date_today = y.ToString() + "/" + m.ToString() + "/" + d.ToString();
            ViewBag.time = date_today;


            List<Models.Sanction> list_letter = context.sanction.Find(x => x.letter_status == "در حال پیگیری" && x.place == R[1] || x.letter_status == "بررسی پاسخ" && x.place == R[1]).ToList();
            return View(list_letter);
        }


        // GET: MO/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MO/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: MO/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(IFormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: MO/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MO/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, IFormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: MO/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: MO/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id, IFormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        public IActionResult export(string type)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            List<Models.Letter> list_letter = new List<Models.Letter>();
            switch (type)
            {
                case "end_time":
                    list_letter = context.Letter.Find(x => x.letter_status == "اتمام مهلت" && x.place == R[1] || x.letter_status == "فوری" && x.place == R[1]).ToList();
                    break;
                case "end":
                    list_letter = context.Letter.Find(x => x.letter_status == "خاتمه" && x.place == R[1]).ToList();
                    break;
                case "in_progress":
                    list_letter = context.Letter.Find(x => x.letter_status == "در حال پیگیری" && x.place == R[1]).ToList();
                    break;
                case "send_H_B":
                    list_letter = context.Letter.Find(x => x.letter_status == "ارسال به بازرسی" && x.place == R[1]).ToList();
                    break;
                default:
                    list_letter = context.Letter.Find(x => x.place == R[1]).ToList();
                    break;
            }
            var file_ex = write_xls(list_letter);

            return file_ex;
        }

        public IActionResult export_sanction(string type)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            List<Models.Sanction> list_letter = new List<Models.Sanction>();
            switch (type)
            {
                case "end_time":
                    list_letter = context.sanction.Find(x => x.letter_status == "اتمام مهلت" && x.place == R[1] || x.letter_status == "فوری" && x.place == R[1]).ToList();
                    break;
                case "end":
                    list_letter = context.sanction.Find(x => x.letter_status == "خاتمه" && x.place == R[1]).ToList();
                    break;
                case "in_progress":
                    list_letter = context.sanction.Find(x => x.letter_status == "در حال پیگیری" && x.place == R[1]).ToList();
                    break;
                case "send_H_B":
                    list_letter = context.sanction.Find(x => x.letter_status == "ارسال به بازرسی" && x.place == R[1]).ToList();
                    break;
                default:
                    list_letter = context.sanction.Find(x => x.place == R[1]).ToList();
                    break;
            }
            var file_ex = write_xls_sanction(list_letter);

            return file_ex;
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult end_letter(string id, string answer_serial, string date_awnser)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }

            List<Models.Letter> letters = context.Letter.Find(x => x.Id == id).ToList();
            Models.Letter lett = letters[0];

            lett.letter_status = "بررسی پاسخ";
            lett.awnser_serial = answer_serial;
            lett.date_awnser = date_awnser;
            context.Letter.DeleteMany(x => x.Id == id);
            context.Letter.InsertOne(lett);

            return Redirect("~/MO_R1/Index_in_progres");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult end_sanction(string id, string answer_serial, string date_awnser)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string ro = chek_Roel();
            List<string> R = ro.Split(" ").ToList();
            //List<string> R = chek_Roel();
            if (R[0] != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }

            List<Models.Sanction> letters = context.sanction.Find(x => x.Id == id).ToList();
            Models.Sanction lett = letters[0];
            lett.letter_status = "بررسی پاسخ";
            lett.awnser_serial = answer_serial;
            lett.date_awnser = date_awnser;
            context.sanction.DeleteMany(x => x.Id == id);
            context.sanction.InsertOne(lett);
            return Redirect("~/MO_R1/Index_in_progres_sanction");
        }


    }
}