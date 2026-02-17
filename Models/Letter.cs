using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace daftar.Models
{
    public class Letter
    {

        [BsonId]
        [BsonRepresentation(MongoDB.Bson.BsonType.ObjectId)]
        public string Id { get; set; }

        [Display(Name = "ارسال کننده")]
        public string sender { get; set; }
        [Display(Name = "سریال نامه")]
        public string serial { get; set; }
        [Display(Name = "عنوان نامه")]
        public string title { get; set; }
        [Display(Name = "شماره ثبت نامه")]
        public string letter_number { get; set; }
        [Display(Name = "تاریخ ایجاد")]
        public string date_create { get; set; }
        [Display(Name = "تاریخ ارجاع")]
        public string date_send { get; set; }
        [Display(Name = "معاونت")]
        public string place { get; set; }
        [Display(Name = "مهلت پیگیری")]
        public string day_reserve { get; set; }
        [Display(Name = "تعداد دفعات پیگیری")]
        public string count_followed { get; set; }
        [Display(Name = "تاریخ پیگیری های انجام شده ")]
        public List<string> list_date_fllowed { get; set; }
        [Display(Name = "مهلت های داده شده")]
        public List<string> last_day_reserve { get; set; }
        //public string last_day_reserve { get; set; }
        [Display(Name = "وضعیت نامه")]
        public string letter_status { get; set; }
        [Display(Name = "سریال پاسخ نامه")]
        public string awnser_serial { get; set; }
        [Display(Name = "تاریخ پاسخ")]
        public string date_awnser { get; set; }
        [Display(Name = "بازه زمانی پاسخ گویی")]
        public string send_until_awnser { get; set; }
        [Display(Name = "تاریخ ارجاع به بازرسی")]
        public string date_send_to_H_B { get; set; }

        [Display(Name = "توضیحات")]
        public string xplain { get; set; }


    }
}
    
	


