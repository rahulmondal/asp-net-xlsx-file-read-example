using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxReadExample.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var filepath = Server.MapPath("~/App_Data/test.xlsx");
            using (var document = SpreadsheetDocument.Open(filepath, false)) 
            {
                var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>();
                ViewBag.NoOfSheets = sheets.Count();
                var sheetName = sheets.Aggregate("", (current, sheet) => current + sheet.Name + ", ");
                ViewBag.SheetName = sheetName;
                
                var relationshipId = sheets.First().Id.Value;
                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

                var firstRow = worksheetPart.Worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>();

                var firstRowFirstCell = firstRow.GetFirstChild<Cell>();

                if (firstRowFirstCell.DataType.Value == CellValues.SharedString)
                {
                    var stringTable = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>()
                        .FirstOrDefault();
                    if (stringTable != null)
                    {
                        ViewBag.FirstRowFirstCell = stringTable.SharedStringTable
                            .ElementAt(int.Parse(firstRowFirstCell.InnerText)).InnerText;
                    }
                    
                }
                else
                {
                    ViewBag.FirstRowFirstCell = firstRowFirstCell.InnerText;
                }
            }
            return View();
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