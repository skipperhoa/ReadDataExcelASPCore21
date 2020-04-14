using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ExcelDataReader;
using ReadExcel.Models;
namespace ReadExcel.Controllers
{
    public class ExcelController : Controller
    {
        private readonly IHostingEnvironment _appEnvironment;

        public ExcelController(IHostingEnvironment appEnvironment)
        {
            this._appEnvironment = appEnvironment;
        }
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Index(IFormCollection collection)
        {
            if (ModelState.IsValid)
            {
                var files = HttpContext.Request.Form.Files;
                List<Product> products = new List<Product>();
                foreach (var item in files)
                {
                    if (item.Length > 0 && item!=null)
                    {
                        string file_name = Guid.NewGuid().ToString().Replace("-", "")+"_"+item.FileName;
                        string uploads = Path.Combine(_appEnvironment.WebRootPath, "files");
                        string urlPart = uploads + "/" + file_name;
                        string extension = Path.GetExtension(urlPart);
                        if(extension ==".xls" || extension == ".xslx")
                        {
                            using (var fileStream = new FileStream(Path.Combine(uploads, file_name), FileMode.Create))
                            {
                                await item.CopyToAsync(fileStream);                                
                            }
                            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                            using (var stream = System.IO.File.Open(urlPart, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    do
                                    {
                                        if (reader.Name == "page1" || reader.Name == "page2")
                                        {
                                            while (reader.Read())
                                            {
                                                products.Add(new Product
                                                {
                                                    idProduct = int.Parse(reader.GetValue(0).ToString()),
                                                    Title = reader.GetValue(1).ToString(),
                                                    UrlImage = reader.GetValue(2).ToString(),
                                                    Price = int.Parse(reader.GetValue(3).ToString()),
                                                    Detail = reader.GetValue(4).ToString()
                                                });
                                            }
                                        }
                                    } while (reader.NextResult());
                                }

                            }

                        }

                    }
                }
                if (products.Count > 0)
                {
                    //return Json(products);
                    ViewBag.products = products;
                    return View();
                }
                return RedirectToAction(nameof(Index));
               
            }
            return RedirectToAction(nameof(Index));
        }
    }
}