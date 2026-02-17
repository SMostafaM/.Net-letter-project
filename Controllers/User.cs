using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using daftar;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using MongoDB.Driver;
using daftar.Models;
using DocumentFormat.OpenXml.Spreadsheet;

namespace daftar.Controllers
{
    public class User : Controller
    {

        private Context context;
        public User()
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

        // GET: User
        public async Task<IActionResult> Index()
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
            try
            {
                List<Models.User> users = new List<Models.User>();
                if (R[1] == "0")
                {
                    users = context.User.Find(_ => true).ToList();
                    ViewBag.user_type = "admin";
                }
                else
                {
                    users = context.User.Find(x => x.role == R[1]).ToList();
                    ViewBag.user_type = "MO";

                }
                //log
                Models.Log log = new Models.Log();
                log.user = get_username();
                log.action = "Get_List";
                log.time = DateTime.Now;
                log.subject = "User";
                log.obj = "all";
                context.Log.InsertOne(log);
                //
                return View(users);
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);

            }
        }

        // GET: User/Details/5
        public async Task<IActionResult> Details(string id)
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
            try
            {
                var user = context.User.Find(x => x.Id == id).ToList();
                Models.User u1 = new Models.User();
                u1 = user[0];
                //log
                Models.Log log = new Models.Log();
                //log.user = Request.Cookies["token"];
                log.action = "Get";
                log.time = DateTime.Now;
                log.subject = "User";
                log.obj = u1.Id;
                context.Log.InsertOne(log);
                //
                return View(u1);
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);

            }
        }

        // GET: User/Create
        public IActionResult Create()
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


            if (R[1] == "0")
            {
                ViewBag.user_type = "admin";
            }
            else
            {
                ViewBag.user_type = "MO";
            }

            return View();
        }

        // POST: User/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create([Bind("Id,user_name,password,token")] Models.User user)
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

            if (user.token != null)
            {
                user.role = user.token;
            }
            try
            {
                security.Coding coding = new security.Coding();
                Models.User user_Chek = new Models.User();
                string S_token = "";
                string S_user = "";

                if (user.user_name != null)
                {
                    user.user_name = user.user_name.Trim().ToLower();
                    List<Models.User> users = context.User.AsQueryable<Models.User>().ToList();
                    user_Chek = users.Find(a => a.user_name == user.user_name);
                }


                if (user_Chek == null)
                {
                    //fill token & sh1 pass
                    if (user.password != null)
                    {
                        string NewPass = coding.Sha1Sum(user.password);
                        user.password = NewPass;
                    }
                    if (R[1] == "0")
                    {
                        ViewBag.user_type = "admin";
                    }
                    else
                    {
                        user.role = R[1] + user.role.ToString();
                        ViewBag.user_type = "MO";
                    }
                    //fill token 
                    /*if (user.role != null)
                    {
                        S_role = coding.Sha1Sum(user.role);
                    }*/
                    //S_role = coding.Sha1Sum(user.role);
                    S_user = coding.Sha1Sum(user.user_name);
                    S_token = coding.Sha1Sum(user.role.ToString());
                    user.token = user.password + S_user + S_token;

                    

                    context.User.InsertOne(user);
                    ViewBag.status = "ok";
                    //log
                    Models.Log log = new Models.Log();
                    log.user = get_username();
                    log.action = "Create_user";
                    log.time = DateTime.Now;
                    log.subject = user.user_name;
                    log.obj = user.Id;
                    context.Log.InsertOne(log);
                    //
                    return View();
                }
                else
                {
                    ViewBag.status = "user_dublicate";
                    return View();
                }
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);

            }

        }

        // GET: User/Edit/5
        public async Task<IActionResult> Edit(string id)
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
            try
            {
                if (R[1] == "0")
                {
                    ViewBag.user_type = "admin";
                }
                else
                {
                    
                    ViewBag.user_type = "MO";
                }

                var user = context.User.Find(x => x.Id == id).ToList();
                Models.User u1 = new Models.User();
                u1 = user[0];
                u1.password = "**********";
                
                return View(u1);
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);

            }
        }

        // POST: User/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(string id, [Bind("Id,user_name,password,token")] Models.User user,string role)
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

            if (role != null){ user.role = role; }
            
            try
            {
                if (R[1] == "0")
                {
                    ViewBag.user_type = "admin";
                }
                else
                {

                    ViewBag.user_type = "MO";
                }

                security.Coding coding = new security.Coding();
                string S_user = "";
                string S_token = "";

                if (string.IsNullOrEmpty(user.password))
                {
                    var use = context.User.Find(x => x.Id == user.Id).ToList();
                    Models.User u1 = new Models.User();
                    u1 = use[0];
                    user.password = u1.password;
                }
                else
                {
                    string NewPass = coding.Sha1Sum(user.password);
                    user.password = NewPass;
                }


                //fill token 
                /*if (user.role != null)
                {
                    S_role = coding.Sha1Sum(user.role);
                }*/
                //#S_role = coding.Sha1Sum(user.role);
                S_user = coding.Sha1Sum(user.user_name);
                S_token = coding.Sha1Sum(user.role);
                user.token = user.password + S_user + S_token;

                context.User.DeleteOne(a => a.Id == user.Id);
                context.User.InsertOne(user);
                //log
                Models.Log log = new Models.Log();
                log.user = get_username();
                log.action = "Edite_user";
                log.time = DateTime.Now;
                log.subject = user.Id;
                log.obj = user.user_name;
                context.Log.InsertOne(log);
                //
                return Redirect("~/User/Index");
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
        }

        // GET: User/Delete/5
        public async Task<IActionResult> Delete(string id)
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
            try
            {
                List<Models.User> letters = context.User.Find(x => x.Id == id).ToList();
                Models.User lett = letters[0];
                //log
                Models.Log log = new Models.Log();
                log.user = get_username();
                log.action = "delete_user";
                log.time = DateTime.Now;
                log.subject = lett.Id;
                log.obj = lett.user_name;
                context.Log.InsertOne(log);
                //
                context.User.DeleteOne(a => a.Id == id);
                
                return Redirect("~/User");
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);

            }
        }


    }
}
