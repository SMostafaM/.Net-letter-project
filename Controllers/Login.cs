using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using daftar.Models;
using MongoDB.Driver;
using System.Globalization;
using System.Diagnostics.Metrics;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace daftar.Controllers
{
    public class Login : Controller
    {
        private Context context;
        public Login()
        {
            context = new Context();
        }

        // GET: Login
        public async Task<IActionResult> login()
        {
            return View();
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
        public string check_letter_time()
        {
            //check letter
            List<Models.Letter> letters_in_progress = new List<Models.Letter>();
            letters_in_progress = context.Letter.Find(x => x.letter_status == "در حال پیگیری").ToList();


            //System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            DateTime dt = DateTime.Now;

            string dat = "";
            int result = 0;
            foreach (var it in letters_in_progress)
            {
                try
                {
                    dat = it.day_reserve;
                    string[] list_le_date = dat.Split("/");
                    PersianCalendar pc = new PersianCalendar();
                    DateTime date_end = new DateTime(int.Parse(list_le_date[0]), int.Parse(list_le_date[1]), int.Parse(list_le_date[2]), pc);
                    //DateTime =it.day_reserve;

                    result = DateTime.Compare(dt, date_end);
                    //if (result < 0)
                    //Console.WriteLine("issue date is less than expired date");
                    //else if (result == 0)
                    //    Console.WriteLine("Both dates are same");
                }
                catch
                {
                    return "line135";
                }

                if (result > 0)
                {
                    it.letter_status = "اتمام مهلت";
                    //string last_res = it.day_reserve;
                    //if (it.last_day_reserve == null)
                    //{
                    //    List<string> l = new List<string>();
                    //    l.Add(last_res);
                    //    it.last_day_reserve = l;
                    //}
                    //else
                    //{
                    //    it.last_day_reserve.Add(last_res);
                    //}
                    //it.day_reserve = "end";
                    context.Letter.DeleteMany(a => a.Id == it.Id);
                    context.Letter.InsertOne(it);
                }


            }

            return "ok";
        }
        public string check_sanction_time()
        {
            //check sanction
            List<Models.Sanction> sanctions_in_progress = new List<Models.Sanction>();
            sanctions_in_progress = context.sanction.Find(x => x.letter_status == "در حال پیگیری").ToList();


            //System.Globalization.PersianCalendar p = new System.Globalization.PersianCalendar();
            DateTime dt = DateTime.Now;

            string dat = "";
            int result = 0;
            foreach (var it in sanctions_in_progress)
            {
                try
                {
                    dat = it.day_reserve;
                    string[] list_le_date = dat.Split("/");
                    PersianCalendar pc = new PersianCalendar();
                    DateTime date_end = new DateTime(int.Parse(list_le_date[0]), int.Parse(list_le_date[1]), int.Parse(list_le_date[2]), pc);
                    //DateTime =it.day_reserve;

                    result = DateTime.Compare(dt, date_end);
                    //if (result < 0)
                    //Console.WriteLine("issue date is less than expired date");
                    //else if (result == 0)
                    //    Console.WriteLine("Both dates are same");
                }
                catch
                {
                    return "line135";
                }

                if (result > 0)
                {
                    it.letter_status = "اتمام مهلت";
                    //string last_res = it.day_reserve;
                    //if (it.last_day_reserve == null)
                    //{
                    //    List<string> l = new List<string>();
                    //    l.Add(last_res);
                    //    it.last_day_reserve = l;
                    //}
                    //else
                    //{
                    //    it.last_day_reserve.Add(last_res);
                    //}
                    //it.day_reserve = "end";
                    context.sanction.DeleteMany(a => a.Id == it.Id);
                    context.sanction.InsertOne(it);
                }


            }

            return "ok";
        }
        public string show_role(string token)
        {
            security.Coding coding = new security.Coding();
            string R = "";
            string Role = "";
            //string cookieValue = Request.Cookies["token"];
            string cookieValue = token;
            for (int i = 80; i < 120; i++)
            {
                Role = Role + cookieValue[i];
            }
            string zerosha1 = coding.func1("0");
            string onesha1 = coding.func1("1");
            if (zerosha1 == Role)
            {
                return "0";
            }
            else if(onesha1==Role)
            {
                return "1";
            }
            else
            {
                return "other";
            }
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> login([Bind("Id,user_name,password,token")] Models.User user)
        {
            if (string.IsNullOrEmpty(user.user_name) || string.IsNullOrEmpty(user.password))
            {
                ViewBag.status = "null";
            }
            else
            {
                try
                {
                    security.Coding coding = new security.Coding();
                    string NewPass = coding.func1(user.password);
                    user.password = NewPass;
                    user.user_name = user.user_name.Trim().ToLower();

                    var user_Chek = context.User.Find(x => x.user_name == user.user_name).ToList();
                    if (user_Chek.Count == 0)
                    {
                        //log
                        Models.Log log = new Models.Log();
                        log.user = get_username();
                        log.action = "login";
                        log.time = DateTime.Now;
                        log.subject = "Bad_username";
                        log.obj = user.user_name;
                        context.Log.InsertOne(log);
                        //
                        ViewBag.status = "NotFound";
                        return View();
                    }
                    else
                    {
                        foreach (var item in user_Chek)
                        {
                            if (item.password == user.password)
                            {
                                //log
                                Models.Log log = new Models.Log();
                                log.user = get_username();
                                log.action = "login";
                                log.time = DateTime.Now;
                                //DateTime date = DateTime.UtcNow;
                                log.subject = "Loged_in";
                                log.obj = item.Id;
                                context.Log.InsertOne(log);
                                //
                                Random random = new Random();
                                // generate a random number
                                int n = random.Next();
                                string n_sha1 = coding.func1(n.ToString());

                                string Active_session = item.token + n_sha1;
                                //insert active sesstion
                                Models.Active_session act=new Models.Active_session();
                                act.user_token = item.token;
                                act.active_sesstion = Active_session;
                                context.Active_session.DeleteMany(a => a.user_token == item.token);
                                context.Active_session.InsertOne(act);

                                var cook = new Microsoft.AspNetCore.Http.CookieOptions() { Path = "/", HttpOnly = false, IsEssential = true, Expires = DateTime.Now.AddHours(10) };
                                Response.Cookies.Append("token", Active_session, cook);
                                string role = show_role(item.token);
                                //var cookieOptions = new CookieOptions();
                                //cookieOptions.Expires = DateTime.Now.AddDays(1);
                                //cookieOptions.Path = "/";
                                //Response.Cookies.Append("SomeCookie", "token", item.token);

                                check_letter_time();
                                check_sanction_time();

                                if (role == "0")
                                {
                                    return Redirect("~/Letter/Index");
                                }
                                else if(role== "other")
                                {
                                    return Redirect("~/MO_R1/Index");
                                }

                            }
                            else
                            {
                                //log
                                Models.Log log = new Models.Log();
                                log.user = get_username();
                                log.action = "login";
                                log.time = DateTime.Now;
                                //DateTime date = DateTime.UtcNow;
                                log.subject = "Bad_Pass";
                                log.obj = item.Id;
                                context.Log.InsertOne(log);
                                //

                                ViewBag.status = "NotFound";
                                return View();
                            }
                        }

                    }

                }
                catch (Exception e)
                {
                    return BadRequest(e.Message);

                }
            }
            return View();
        }



        // GET: Login/Details/5
        public async Task<IActionResult> Details(string id)
        {
           
            return View();
        }

        // GET: Login/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Login/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,user,time,action,subject,obj")] Log log)
        {
            
            return View();
        }

        // GET: Login/Edit/5
        public async Task<IActionResult> Edit(string id)
        {
            
            return View();
        }

        // POST: Login/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(string id, [Bind("Id,user,time,action,subject,obj")] Log log)
        {
            return View();
        }

        public IActionResult exit()
        {
            Response.Cookies.Delete("token");
            return Redirect("~/login/login");


        }

    }
}
