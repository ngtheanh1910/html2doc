using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Web.Http;

namespace DemoHtml2Doc.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public void Export2Word()
        {
            string _fileCSS = Server.MapPath("~/css/style.css");
            string _strCSS = System.IO.File.ReadAllText(_fileCSS);
            string _baseURL = "http://localhost:1385/";
            string _filename =  "document.docx";
            string htmlRaw = @"<div id='todoapp'><div id='header'><input type='text' id='new-todo' name='new-todo' placeholder='Enter some task'><button id='btnAdd'>Add</button></div><div id='main'><ul id='todo-list'><li><span>Nguyễn Tuấn Anh</span><input class='edit' value='Nguyễn Tuấn Anh'><div class='destroy'><i class='far fa-trash-alt'></i></div></li><li><span>Nguyễn Quang Hải</span><input class='edit' value='Nguyễn Quang Hải'><div class='destroy'><i class='far fa-trash-alt'></i></div></li><li><span>Nguyễn Công Phượng</span><input class='edit' value='Nguyễn Công Phượng'><div class='destroy'><i class='far fa-trash-alt'></i></div></li><li><span>hello</span><input class='edit' value='hello'><div class='destroy'><i class='far fa-trash-alt'></i></div></li><li><span>Nguyễn Hoàng Anh</span><input class='edit' value='Nguyễn Hoàng Anh'><div class='destroy'><i class='far fa-trash-alt'></i></div></li></ul></div></div>";

            StringBuilder strHTML = new StringBuilder();
            strHTML.Append("<html " +
                " xmlns:o='urn:schemas-microsoft-com:office:office'" +
                " xmlns:w='urn:schemas-microsoft-com:office:word'" +
                " xmlns='http://www.w3.org/TR/REC-html40'>" +
                "<head><title>Invoice Sample</title>");

            strHTML.Append("<xml><w:WordDocument>" +
                " <w:View>Print</w:View>" +
                " <w:Zoom>100</w:Zoom>" +
                " <w:DoNotOptimizeForBrowser/>" +
                " </w:WordDocument>" +
                " </xml>");

            strHTML.Append("<style>" + _strCSS + "</style></head>");
            strHTML.Append("<body><div class='page-settings'>" + htmlRaw + "</div></body></html>");

            Response.AppendHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            Response.AppendHeader("Content-disposition", "attachment;filename=" + _filename + "");
            Response.Write(strHTML.ToString());
        }

        public JsonResult Export2WordAjax(string html)
        {
            string _fileCSS = Server.MapPath("~/css/style.css");
            string _strCSS = System.IO.File.ReadAllText(_fileCSS);
            string _baseURL = "http://localhost:1385/";
            string _filename = "document.docx";
            string htmlRaw = html.ToString();

            StringBuilder strHTML = new StringBuilder();
            strHTML.Append("<html " +
                " xmlns:o='urn:schemas-microsoft-com:office:office'" +
                " xmlns:w='urn:schemas-microsoft-com:office:word'" +
                " xmlns='http://www.w3.org/TR/REC-html40'>" +
                "<head><title>Invoice Sample</title>");

            strHTML.Append("<xml><w:WordDocument>" +
                " <w:View>Print</w:View>" +
                " <w:Zoom>100</w:Zoom>" +
                " <w:DoNotOptimizeForBrowser/>" +
                " </w:WordDocument>" +
                " </xml>");

            strHTML.Append("<style>" + _strCSS + "</style></head>");
            strHTML.Append("<body><div class='page-settings'>" + htmlRaw + "</div></body></html>");

            Response.AppendHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            Response.AppendHeader("Content-disposition", "attachment;filename=" + _filename + "");
            Response.Write(strHTML.ToString());
            return Json(null, JsonRequestBehavior.AllowGet);
        }
        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}