using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace HuiQianDan.Controllers
{
    public class SetDataRegAndSealController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public SetDataRegAndSealController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            string userName = Request.Query["userName"];
            //***************************卓正PageOffice组件的使用********************************
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();

            PageOfficeNetCore.WordWriter.DataRegion d1 = doc.OpenDataRegion("PO_1");
            PageOfficeNetCore.WordWriter.DataRegion d2 = doc.OpenDataRegion("PO_2");
            PageOfficeNetCore.WordWriter.DataRegion d3 = doc.OpenDataRegion("PO_3");

            //设置数据区域文本样式
            d1.Font.Color = Color.Green;
            d2.Font.Color = Color.Blue;
            d3.Font.Color = Color.Magenta;

            //根据登录用户名设置数据区域可编辑性
            //张三登录后
            if (userName.Equals("zhangsan"))
            {
                userName = "张三";
                d1.Editing = true;
                d2.Editing = false;
                d3.Editing = false;
            }
            //李四登录后
            else if (userName.Equals("lisi"))
            {
                userName = "李四";
                d1.Editing = false;
                d2.Editing = true;
                d3.Editing = false;
            }
            //王五登录后
            else
            {
                userName = "王五";
                d1.Editing = false;
                d2.Editing = false;
                d3.Editing = true;
            }

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.AddCustomToolButton("-", "", 0);
            pageofficeCtrl.AddCustomToolButton("盖章方式一", "AddSeal", 2);
            pageofficeCtrl.AddCustomToolButton("盖章方式二", "AddSeal2", 2);
            pageofficeCtrl.AddCustomToolButton("-", "", 0);
            pageofficeCtrl.AddCustomToolButton("签字方式一", "AddSign", 3);
            pageofficeCtrl.AddCustomToolButton("签字方式二", "AddSign2", 3);
            pageofficeCtrl.AddCustomToolButton("-", "", 0);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            pageofficeCtrl.SetWriter(doc);
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.userName = userName;
            return View();
        }
        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/doc/" + fs.FileName);
            fs.Close();
            return Content("OK");
        }
    }
}