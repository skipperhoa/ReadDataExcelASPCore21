# Read Data Excel ASP.NET Core 2.1
**Step 1:** Open Visual Studio 2019 -> File->New-> Create Project

**Step 2:** Select .NET ASP.NET Core Web Application template

**Step 3:** Enter Project Name, Select .NET Core, ASP.NET Core 2.1, Chọn Web Application (Model-View-Controller), Create

**Step 4:** Open Nuget Packager Manager -> Install Plugin 

**ExcelDataReader**

**System.Text.Encoding.CodePages**

[ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)

**Step 5:** Create Product.cs file in directory Models

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReadExcel.Models
{
    public class Product
    {
        public int idProduct { get; set; }
        public string Title { get; set; }
        public string UrlImage { get; set; }
        public string Detail { get; set; }
        public int Price { get; set; }


    }
}
```

**Step 6**: Create ExcelControlelr file in directory Controllers, after then open it, change the following code below
```csharp
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
```
**Step 7:** Create Index.cshtml file in Views/Excel directory
```csharp
@model IEnumerable<Product>
@{
    ViewData["Title"] = "Index";
    Layout = "~/Views/Shared/_Layout.cshtml";
}
<div class="row">
    <div class="col-md-8">
        <h2>Read file</h2>
        <form asp-controller="Excel" asp-action="Index" enctype="multipart/form-data">
           @Html.AntiForgeryToken()
            <div class="form-group">
                <label>Choose excel file</label>
                <input type="file" name="file" class="form-control"/>
            </div>
            <div class="form-group">
                <input type="submit" name="submit" value="Upload"/>
            </div>
        </form>
    </div>
    <div class="col-md-8">
        @if(ViewBag.products!=null)
        {
            <table class="table">
                <tr>
                    <th>ID</th>
                    <th>Title</th>
                    <th>UrlImage</th>
                    <th>Price</th>
                    <th>Detail</th>
                </tr>
                @foreach(var item in ViewBag.products)
                {
                    <tr>
                        <td>@item.idProduct</td>
                        <td>@item.Title</td>
                        <td>@item.UrlImage</td>
                        <td>@item.Price</td>
                        <td>@item.Detail</td>
                    </tr>
                }
            </table>
        }
    </div>
</div>

```

