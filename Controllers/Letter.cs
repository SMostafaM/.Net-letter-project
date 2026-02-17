using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using daftar.Models;
using MongoDB.Driver;
using MongoDB.Bson;
using static System.Runtime.InteropServices.JavaScript.JSType;

using System.IO;
using System.Data;
using ClosedXML.Excel;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics.Metrics;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.EntityFrameworkCore.Query.SqlExpressions;

namespace daftar.Controllers
{
    public class Letter : Controller
    {
        private Context context;
        public Letter()
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
                    if(acs[0].active_sesstion== cookieValue)
                    {
                        var cook = new Microsoft.AspNetCore.Http.CookieOptions() { Path = "/", HttpOnly = false, IsEssential = true, Expires = DateTime.Now.AddHours(10) };
                        Response.Cookies.Append("token", cookieValue, cook);
                        C = 1;
                    }
                }
               
            }

            return C;
        }
        public string get_username()
        {
            string username = "";
            try
            {
                username = User.Identity.Name.ToString();
            }
            catch (Exception e)
            {
                username = e.ToString();
            }
            //try
            //{
            //    username = username.Substring(username.Length - 8);
            //}
            //catch
            //{
            //    username = "error in 8 end";
            //}
            return username;
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
            string zerosha1 = coding.Sha1Sum("0");
            if (zerosha1 == Role)
            {
                return "1";
            }
            else
            {
                return "0";
            }
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

                worksheet.Cell(currentRow, 1).Value = "ارسال کننده";
                worksheet.Cell(currentRow, 2).Value = "سریال نامه";
                worksheet.Cell(currentRow, 3).Value = "عنوان نامه";
                worksheet.Cell(currentRow, 4).Value = "شماره نامه وارده";
                worksheet.Cell(currentRow, 5).Value = "تاریخ ایجاد نامه";
                worksheet.Cell(currentRow, 6).Value = "تاریخ ارجاع به معاونت مربوطه";
                worksheet.Cell(currentRow, 7).Value = "معاونت";
                worksheet.Cell(currentRow, 8).Value = "مهلت پاسخ";
                worksheet.Cell(currentRow, 9).Value = "دفعات پیگیری";
                worksheet.Cell(currentRow, 10).Value = "مهلت های داده شده";
                worksheet.Cell(currentRow, 11).Value = "تاریخ پیگیری های انجام شده";
                worksheet.Cell(currentRow, 12).Value = "وضعیت نامه";
                worksheet.Cell(currentRow, 13).Value = "سریال پاسخ نامه";
                worksheet.Cell(currentRow, 14).Value = "تاریخ پاسخ نامه";
                worksheet.Cell(currentRow, 15).Value = "مدت پاسخ دهی";
                worksheet.Cell(currentRow, 16).Value = "تاریخ به بارزسی";
                worksheet.Cell(currentRow, 17).Value = "توضیحات";



                string last_d = "";
                foreach (var item in list_to_write)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = item.sender;
                    worksheet.Cell(currentRow, 2).Value = item.serial;
                    worksheet.Cell(currentRow, 3).Value = item.title;
                    worksheet.Cell(currentRow, 4).Value = item.letter_number;
                    worksheet.Cell(currentRow, 5).Value = item.date_create;
                    worksheet.Cell(currentRow, 6).Value = item.date_send;
                    worksheet.Cell(currentRow, 7).Value = item.place;
                    worksheet.Cell(currentRow, 8).Value = item.day_reserve;
                    worksheet.Cell(currentRow, 9).Value = item.count_followed;
                    if (item.last_day_reserve != null)
                    {
                        foreach (string i in item.last_day_reserve)
                        {
                            last_d = last_d + i.ToString();
                        }
                    }
                    worksheet.Cell(currentRow, 10).Value = last_d;
                    last_d = "";
                    if (item.list_date_fllowed != null)
                    {
                        foreach (string i in item.list_date_fllowed)
                        {
                            last_d = last_d + i.ToString();
                        }
                    }
                    worksheet.Cell(currentRow, 11).Value = last_d;
                    worksheet.Cell(currentRow, 12).Value = item.letter_status;
                    worksheet.Cell(currentRow, 13).Value = item.awnser_serial;
                    worksheet.Cell(currentRow, 14).Value = item.date_awnser;
                    worksheet.Cell(currentRow, 15).Value = item.send_until_awnser;
                    worksheet.Cell(currentRow, 16).Value = item.date_send_to_H_B;
                    worksheet.Cell(currentRow, 17).Value = item.xplain;

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

        public string moon_len(int i, string y)
        {
            string date_today = "";
            if (i < 10)
            {
                date_today = y.ToString() + "/0" + i.ToString();
            }
            else
            {
                date_today = y.ToString() + "/" + i.ToString();
            }
            return date_today;
        }
        public List<Models.Letter> search_letter(string y, int m1, int m2, string letter_type = "خاتمه", string search_date = "awnser", string mo = "all")
        {
            //solar date
            System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            DateTime dt = DateTime.Now;
            int y_now = p.GetYear(dt);
            //int m = p.GetMonth(dt);
            //int d = p.GetDayOfMonth(dt);
            string date_today = "";
            ViewBag.year = y_now - 5;

            ViewBag.Dyear = y;
            ViewBag.moon = m1;
            //ViewBag.now_page = follow_page;
            List<Models.Letter> list_letter = new List<Models.Letter>();
            //List<Models.Letter> list_letter_total=new List<Models.Letter>();

            if (m1 == 50 || m2 == 50)
            {

                date_today = y.ToString();
                if (search_date == "awnser")
                {
                    if (mo != "all")
                    {
                        if (letter_type == "all")
                        {
                            list_letter = context.Letter.Find(x => x.date_awnser.Contains(date_today) && x.place == mo).ToList();
                        }
                        else
                        {
                            list_letter = context.Letter.Find(x => x.letter_status == letter_type && x.date_awnser.Contains(date_today) && x.place == mo).ToList();
                        }
                    }
                    else
                    {
                        if (letter_type == "all")
                        {
                            list_letter = context.Letter.Find(x => x.date_awnser.Contains(date_today)).ToList();
                        }
                        else
                        {
                            list_letter = context.Letter.Find(x => x.letter_status == letter_type && x.date_awnser.Contains(date_today)).ToList();
                        }
                        
                    }
                }
                else
                {

                    if (mo != "all")
                    {
                        if (letter_type == "all")
                        {
                            list_letter = context.Letter.Find(x => x.date_send.Contains(date_today) && x.place == mo).ToList();
                        }
                        else
                        {
                            list_letter = context.Letter.Find(x => x.letter_status == letter_type && x.date_send.Contains(date_today) && x.place == mo).ToList();
                        }
                    }
                    else
                    {
                        if (letter_type == "all")
                        {
                            list_letter = context.Letter.Find(x => x.date_send.Contains(date_today)).ToList();
                        }
                        else
                        {
                            list_letter = context.Letter.Find(x => x.letter_status == letter_type && x.date_send.Contains(date_today)).ToList();
                        }

                    }
                }
            }
            else
            {
                if (m1 > m2)
                {
                    ViewBag.date = "false";
                    return (list_letter);
                }


                if (search_date == "awnser")
                {
                    if (mo != "all")
                    {
                        if (letter_type == "all")
                        {
                            for (int i = m1; i <= m2; i++)
                            {
                                date_today = moon_len(i, y);
                                list_letter = list_letter.Concat(context.Letter.Find(x => x.date_awnser.Contains(date_today) && x.place == mo).ToList()).ToList();
                            }
                        }
                        else
                        {
                            for (int i = m1; i <= m2; i++)
                            {
                                date_today = moon_len(i, y);
                                list_letter = list_letter.Concat(context.Letter.Find(x => x.letter_status == letter_type && x.date_awnser.Contains(date_today) && x.place == mo).ToList()).ToList();
                            }
                        }
                    }
                    else
                    {
                        if (letter_type == "all")
                        {
                            for (int i = m1; i <= m2; i++)
                            {
                                date_today = moon_len(i, y);
                                list_letter = list_letter.Concat(context.Letter.Find(x => x.date_awnser.Contains(date_today)).ToList()).ToList();

                            }
                        }
                        else
                        {
                            for (int i = m1; i <= m2; i++)
                            {
                                date_today = moon_len(i, y);
                                list_letter = list_letter.Concat(context.Letter.Find(x => x.letter_status == letter_type && x.date_awnser.Contains(date_today)).ToList()).ToList();
                            }
                        }

                    }
                }
                else
                {

                    if (mo != "all")
                    {
                        if (letter_type == "all")
                        {
                            for (int i = m1; i <= m2; i++)
                            {
                                date_today = moon_len(i, y);
                                list_letter = list_letter.Concat(context.Letter.Find(x => x.date_send.Contains(date_today) && x.place == mo).ToList()).ToList();
                            }
                        }
                        else
                        {
                            for (int i = m1; i <= m2; i++)
                            {
                                date_today = moon_len(i, y);
                                list_letter = list_letter.Concat(context.Letter.Find(x => x.letter_status == letter_type && x.date_send.Contains(date_today) && x.place == mo).ToList()).ToList();
                            }
                        }
                    }
                    else
                    {
                        if (letter_type == "all")
                        {
                            for (int i = m1; i <= m2; i++)
                            {
                                date_today = moon_len(i, y);
                                list_letter = list_letter.Concat(context.Letter.Find(x => x.date_send.Contains(date_today)).ToList()).ToList();
                            }
                        }
                        else
                        {
                            for (int i = m1; i <= m2; i++)
                            {
                                date_today = moon_len(i, y);
                                list_letter = list_letter.Concat(context.Letter.Find(x => x.letter_status == letter_type && x.date_send.Contains(date_today)).ToList()).ToList();
                            }
                        }

                    }
                }

            }
            
            //list_letter = context.Letter.Find(x => x.letter_status == "خاتمه" && x.date_awnser.Contains(date_today)).ToList();
            return (list_letter);
        }


        // GET: Letter
        public IActionResult Index(int follow_page=1)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
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

            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "اتمام مهلت";
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = "اتمام مهلت";
            context.Log.InsertOne(log);
            //
            
            List<Models.Letter> list_letter = context.Letter.Find(x=>x.letter_status=="اتمام مهلت" || x.letter_status == "فوری"||x.letter_status== "بررسی پاسخ").ToList();
           
            int count_val=list_letter.Count();
            int count_page = count_val / 30;
            ViewBag.count_page = count_page + 1;
            if (follow_page == -5)
            {
                ViewBag.now_page = -5;
                return View(list_letter);
            }
            if (follow_page <= 0)
            {
                follow_page = 1;
            }
            ViewBag.now_page = follow_page;
            list_letter = context.Letter.Find(x=>x.letter_status=="اتمام مهلت" || x.letter_status == "فوری"||x.letter_status== "بررسی پاسخ").Skip((follow_page - 1) * 30).Limit(30).ToList();
            return View(list_letter);
        }
        // GET: Letter
        public IActionResult Index_end(int year = 0, int month = 0,string letter_type= "خاتمه")
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            //solar date
            System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            DateTime dt = DateTime.Now;
            int y = p.GetYear(dt);
            int m = p.GetMonth(dt);
            //int d = p.GetDayOfMonth(dt);
            string date_today = "";
            if (m < 10)
            {
                 date_today = y.ToString() + "/0" + m.ToString();
            }
            else
            {
                 date_today = y.ToString() + "/" + m.ToString();
            }
            ViewBag.year = y - 5;
            ViewBag.Dyear = y;
            ViewBag.moon = m;
            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "خاتمه";
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = date_today;
            context.Log.InsertOne(log);
            //
            //if (year == 0 && month == 0)
            //{
            //    ViewBag.Dyear = y;
            //    ViewBag.moon = m;
            //}
            //else
            //{
            //    ViewBag.Dyear = year;
            //    ViewBag.moon = month;
            //    if (m < 10)
            //    {
            //        date_today = y.ToString() + "/0" + m.ToString();
            //    }
            //    else
            //    {
            //        date_today = y.ToString() + "/" + m.ToString();
            //    }
            //}
            List<Models.Letter> list_letter = context.Letter.Find(x => x.letter_status == "خاتمه" && x.date_awnser.Contains(date_today)).ToList();

            //int count_val = list_letter.Count();
            //int count_page = count_val / 30;
            //ViewBag.count_page = count_page + 1;
            //if (follow_page == -5)
            //{
            //    ViewBag.now_page = -5;
            //    return View(list_letter);
            //}
            //if (follow_page <= 0)
            //{
            //    follow_page = 1;
            //}
            //ViewBag.now_page = follow_page;
            //list_letter = context.Letter.Find(x => x.letter_status == "خاتمه" && x.date_awnser.Contains(date_today)).Skip((follow_page - 1) * 30).Limit(30).ToList();
            return View(list_letter);
        }
        [HttpPost]
        // GET: Letter
        public IActionResult Index_end(string y,int m1, int m2, int follow_page=1, string letter_type = "خاتمه",string search_date= "awnser",string mo="all",string let_num="0",string type_num= "serial")
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }


            List<Models.Letter> list_letter=new List<Models.Letter>();
            if (let_num == "0")
            {
                list_letter = search_letter(y, m1, m2, letter_type, search_date, mo);
            }
            else
            {
                switch (type_num)
                {
                    case "serial":
                        list_letter = context.Letter.Find(x => x.serial == let_num).ToList();
                        break;
                    case "letter_number":
                        list_letter = context.Letter.Find(x => x.letter_number == let_num).ToList();
                        break;
                }
                        
                //solar date
                System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
                DateTime dt = DateTime.Now;
                int y_now = p.GetYear(dt);
                //int m = p.GetMonth(dt);
                //int d = p.GetDayOfMonth(dt);
                string date_today = "";
                ViewBag.year = y_now - 5;

                ViewBag.Dyear = y;
                ViewBag.moon = m1;
            }



            ViewBag.y = y;
            ViewBag.m1=m1;
            ViewBag.m2= m2;
            ViewBag.letter_type = letter_type;
            ViewBag.search_date = search_date;
            ViewBag.mo = mo;

            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = letter_type;
            log.time = DateTime.Now;
            log.subject = "letter_serach"+ mo;
            log.obj = y.ToString()+m1.ToString()+m2.ToString();
            context.Log.InsertOne(log);
            //
            //list_letter = context.Letter.Find(x => x.letter_status == "خاتمه" && x.date_awnser.Contains(date_today)).ToList();
            return View(list_letter);
            }




        public IActionResult Index_send_H_B(int follow_page = 1)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
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

            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "ارسال به بازرسی";
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = "ارسال به بازرسی";
            context.Log.InsertOne(log);
            //
            List<Models.Letter> list_letter = context.Letter.Find(x => x.letter_status == "ارسال به بازرسی").ToList();
            int count_val = list_letter.Count();
            int count_page = count_val / 30;
            ViewBag.count_page = count_page + 1;
            if (follow_page == -5)
            {
                ViewBag.now_page = -5;
                return View(list_letter);
            }
            if (follow_page <= 0)
            {
                follow_page = 1;
            }
            ViewBag.now_page = follow_page;
            list_letter = context.Letter.Find(x => x.letter_status == "ارسال به بازرسی").Skip((follow_page - 1) * 30).Limit(30).ToList();
            return View(list_letter);
        }
        public IActionResult Index_in_progres(int follow_page = 1)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
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
            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "در حال پیگیری";
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = "در حال پیگیری";
            context.Log.InsertOne(log);
            //
            List<Models.Letter> list_letter = context.Letter.Find(x => x.letter_status == "در حال پیگیری").ToList();

            int count_val = list_letter.Count();
            int count_page = count_val / 30;
            ViewBag.count_page = count_page + 1;
            if (follow_page == -5)
            {
                ViewBag.now_page = -5;
                return View(list_letter);
            }
            if (follow_page <= 0)
            {
                follow_page = 1;
            }
            ViewBag.now_page = follow_page;
            list_letter = context.Letter.Find(x => x.letter_status == "در حال پیگیری").Skip((follow_page - 1) * 30).Limit(30).ToList();
            return View(list_letter);
        }

        // GET: Letter/Details/5
        public async Task<IActionResult> Details(string id)
        {
           

            return View();
        }

        // GET: Letter/Create
        public IActionResult Create() {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
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

            dt = dt.AddDays(7);
            y = p.GetYear(dt);
            m = p.GetMonth(dt);
            d = p.GetDayOfMonth(dt);
            string date_res = y.ToString() + "/" + m.ToString() + "/" + d.ToString();
            ViewBag.time_res = date_res;

            return View();
        }

        // POST: Letter/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,serial,title,letter_number,date_create,date_send,place,day_reserve,count_followed,last_day_reserve,letter_status,awnser_serial,sender,xplain")] Models.Letter letter)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
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

            dt = dt.AddDays(7);
            y = p.GetYear(dt);
            m = p.GetMonth(dt);
            d = p.GetDayOfMonth(dt);
            string date_res = y.ToString() + "/" + m.ToString() + "/" + d.ToString();
            ViewBag.time_res = date_res;


            var list_let = context.Letter.Find(x => x.serial == letter.serial && x.date_send == letter.date_send && x.place==letter.place).ToList();
            if (list_let.Count() == 0)
            {
                letter.count_followed = "1";
                letter.letter_status = "در حال پیگیری";
                if (letter.day_reserve == date_today)
                {
                    letter.letter_status = "فوری";
                }
                //log
                Models.Log log = new Models.Log();
                log.user = get_username();
                log.action = "create";
                log.time = DateTime.Now;
                log.subject = "letter";
                log.obj = letter.serial;
                context.Log.InsertOne(log);
                //
                context.Letter.InsertOne(letter);
                ViewBag.status = "ok";
            }
            else
            {
                ViewBag.status = "exist";
            }


            return View();
        }

        // GET: Letter/Edit/5
        public async Task<IActionResult> Edit(string id)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            try
            {
                List<Models.Letter> letters = context.Letter.Find(x => x.Id == id).ToList();
                Models.Letter lett = letters[0];
                ViewBag.letter = lett;
                return View();
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
            return View();
        }

        // POST: Letter/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(string id, [Bind("Id,serial,title,letter_number,date_create,date_send,place,day_reserve,count_followed,last_day_reserve,letter_status,awnser_serial,sender,xplain")] Models.Letter letter)
        {
            try
            {
                int C = Chek_cooke();
                if (C != 1)
                {
                    return Redirect("~/login/login");
                }
                string R = chek_Roel();
                if (R != "1")
                {
                    Response.Cookies.Delete("token");
                    return Redirect("~/login/login");
                }
                    List<Models.Letter> letters = context.Letter.Find(x => x.Id == id).ToList();
                Models.Letter letter_old = letters[0];
                if (letter.letter_status != null)
                {
                    letter_old.letter_status = letter.letter_status;
                }
                letter_old.serial = letter.serial;
                letter_old.letter_number = letter.letter_number;
                letter_old.title = letter.title;
                letter_old.date_create = letter.date_create;
                letter_old.date_send = letter.date_send;
                letter_old.place = letter.place;
                letter_old.day_reserve = letter.day_reserve;

                //log
                Models.Log log = new Models.Log();
                log.user = get_username();
                log.action = "edite";
                log.time = DateTime.Now;
                log.subject = "letter";
                log.obj = letter.serial;
                context.Log.InsertOne(log);
                //
                context.Letter.DeleteMany(a => a.Id == id);
                context.Letter.InsertOne(letter);
                //log
                //Models.Log log = new Models.Log();
                //log.user = Request.Cookies["token"];
                //log.action = "Edite";
                //log.time = DateTime.Now;
                //log.subject = "home_data";
                //log.obj = home_data.Id;
                //context.Log.InsertOne(log);
                //
                return Redirect("~/Letter/Index");
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }

            //return View();
        }

        // GET: Letter/Delete/5
        public IActionResult Delete(string id)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            List<Models.Letter> letters = context.Letter.Find(x => x.Id == id).ToList();
            context.deleted_Letter.InsertMany(letters);
            Models.Letter lett = letters[0];
            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "delete";
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = lett.serial;
            context.Log.InsertOne(log);
            //
            context.Letter.DeleteMany(x => x.Id == id);

            string v = HttpContext.Request.Headers.Referer.ToString();
            return Redirect(v);
        }

        // POST: Letter/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(string id)
        {

            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult end_letter(string id,string answer_serial,string date_awnser)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            
            List<Models.Letter> letters = context.Letter.Find(x => x.Id == id).ToList();
            Models.Letter lett = letters[0];
            lett.letter_status = "خاتمه";
            lett.awnser_serial = answer_serial;
            lett.date_awnser = date_awnser;

            string start = lett.date_send;
            string end =  lett.date_awnser;

            
            string[] list_le_date = start.Split("/");
            PersianCalendar pc = new PersianCalendar();
            DateTime date_start = new DateTime(int.Parse(list_le_date[0]), int.Parse(list_le_date[1]), int.Parse(list_le_date[2]), pc);
            //DateTime =it.day_reserve;
            list_le_date = end.Split("/");
            PersianCalendar pc2 = new PersianCalendar();
            DateTime date_end = new DateTime(int.Parse(list_le_date[0]), int.Parse(list_le_date[1]), int.Parse(list_le_date[2]), pc2);

            //var result = DateTime.Compare(date_start, date_end);

            var diffOfDates = date_end - date_start;
            lett.send_until_awnser = diffOfDates.Days.ToString();

            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "end_letter";
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = lett.serial;
            context.Log.InsertOne(log);
            //

            context.Letter.DeleteMany(x => x.Id == id);
            context.Letter.InsertOne(lett);

            string v = HttpContext.Request.Headers.Referer.ToString();
            return Redirect(v);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult send_H_B(string id, string date_send_H_B)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }

            List<Models.Letter> letters = context.Letter.Find(x => x.Id == id).ToList();
            Models.Letter lett = letters[0];
            lett.letter_status = "ارسال به بازرسی";
            lett.date_send_to_H_B = date_send_H_B;


            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "send_H_B";
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = lett.serial;
            context.Log.InsertOne(log);
            //
            context.Letter.DeleteMany(x => x.Id == id);
            context.Letter.InsertOne(lett);

            string v = HttpContext.Request.Headers.Referer.ToString();
            return Redirect(v);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult continuation(string id,string day,string date_fllowed)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            List<Models.Letter> let_lis= context.Letter.Find(x => x.Id == id).ToList();
            Models.Letter let = let_lis[0];
            int c = int.Parse(let.count_followed) + 1;
            let.count_followed = c.ToString();
            let.letter_status = "در حال پیگیری";

            if (let.last_day_reserve == null)
            {
                List<string> last_dat = new List<string>();
                last_dat.Add(let.day_reserve);
                let.last_day_reserve = last_dat;
            }
            else
            {
                let.last_day_reserve.Add(let.day_reserve);
            }
            if (let.list_date_fllowed == null)
            {
                List<string> last_date_f = new List<string>();
                last_date_f.Add(date_fllowed);
                let.list_date_fllowed = last_date_f;
            }
            else
            {
                let.list_date_fllowed.Add(date_fllowed);
            }


            let.day_reserve = day;



            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            string ac = "continuation" + let.day_reserve.ToString();
            log.action = ac;
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = let.serial;
            context.Log.InsertOne(log);
            //
            context.Letter.DeleteMany(x => x.Id == id);
            context.Letter.InsertOne(let);



            return Redirect("~/Letter/Index");
        }

        public IActionResult export(string type,string y,string moon, int m2=0, string letter_type = "خاتمه", string search_date = "awnser", string mo = "all")
        {
            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "export:"+y+","+ moon+", month to:"+ m2.ToString();
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = type;
            context.Log.InsertOne(log);
            //
            List<Models.Letter> list_letter = new List<Models.Letter>();
            switch (type)
            {
                case "end_moon":
                    int m1=int.Parse(moon);
                    list_letter = search_letter(y, m1, m2, letter_type, search_date, mo);

                    break;
                case "end_time":
                    list_letter = context.Letter.Find(x => x.letter_status == "اتمام مهلت" || x.letter_status == "فوری").ToList();
                    break;
                case "end":
                    list_letter = context.Letter.Find(x => x.letter_status == "خاتمه").ToList();
                    break;
                case "in_progress":
                    list_letter = context.Letter.Find(x => x.letter_status == "در حال پیگیری").ToList();
                    break;
                case "send_H_B":
                    list_letter = context.Letter.Find(x => x.letter_status == "ارسال به بازرسی").ToList();
                    break;
                default:
                    list_letter = context.Letter.Find(x => true).ToList();
                    break;
            }
            var file_ex=write_xls(list_letter);

            return file_ex;
        }
        [HttpPost]
        public IActionResult xplain(string xplain,string id)
        {
            int C = Chek_cooke();
            if (C != 1)
            {
                return Redirect("~/login/login");
            }
            string R = chek_Roel();
            if (R != "1")
            {
                Response.Cookies.Delete("token");
                return Redirect("~/login/login");
            }
            List<Models.Letter> le = context.Letter.Find(x => x.Id == id).ToList();
            Models.Letter letter = le[0];
            //letter.xplain = letter.xplain + "/n" + xplain;
            letter.xplain = xplain;

            //log
            Models.Log log = new Models.Log();
            log.user = get_username();
            log.action = "xplain";
            log.time = DateTime.Now;
            log.subject = "letter";
            log.obj = letter.serial+":"+letter.xplain;
            context.Log.InsertOne(log);
            //
            context.Letter.DeleteOne(x => x.Id == id);
            context.Letter.InsertOne(letter);


            string v = HttpContext.Request.Headers.Referer.ToString();
            return Redirect(v);
        }

    }
}
 