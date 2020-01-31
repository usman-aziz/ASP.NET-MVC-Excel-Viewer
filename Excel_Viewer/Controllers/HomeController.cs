using Aspose.Cells;
using Aspose.Cells.Rendering;
using Excel_Viewer.Helper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Excel_Viewer.Controllers
{
    public class HomeController : Controller
    {
        public List<Sheet> sheets;

        [HttpGet]
        public ActionResult Index(string fileName)
        {
            sheets = new List<Sheet>();
            if (fileName == null)
            {
                // Display default Worksheet on page load
                sheets = RenderExcelWorksheetsAsImage("Workbook.xlsx");
            }
            else
            {
                sheets = RenderExcelWorksheetsAsImage(fileName);
            }

            return View(sheets);
        }
        public List<Sheet> RenderExcelWorksheetsAsImage(string FileName)
        {
            // Load the Excel workbook 
            Workbook book = new Workbook(Server.MapPath(Path.Combine("~/Documents", FileName)));
            var workSheets = new List<Sheet>();
            // Set image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.HorizontalResolution = 200;
            options.VerticalResolution = 200;
            options.AllColumnsInOnePagePerSheet = true;
            options.OnePagePerSheet = true;
            options.TextCrossType = TextCrossType.Default;
            options.ImageType = Aspose.Cells.Drawing.ImageType.Png;

            string imagePath = "";
            string basePath = Server.MapPath("~/");

            // Create Excel workbook renderer
            WorkbookRender wr = new WorkbookRender(book, options);
            // Save and view worksheets
            for (int j = 0; j < book.Worksheets.Count; j++)
            {
                imagePath = Path.Combine("/Documents/Rendered", string.Format("sheet_{0}.png", j));
                wr.ToImage(j, basePath + imagePath);
                workSheets.Add(new Sheet { SheetName = string.Format("{0}", book.Worksheets[j].Name), Path = imagePath });
            }

            return workSheets;
        }
    }
}