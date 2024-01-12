using ImportarPlanilhaExcelProjeto.Data;
using ImportarPlanilhaExcelProjeto.Models;
using ImportarPlanilhaExcelProjeto.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;

namespace ImportarPlanilhaExcelProjeto.Controllers
{
    public class HomeController : Controller
    {
        private readonly AppDbContext _context;
        private readonly IExcelInterface _excelInterface;

        public HomeController(AppDbContext context, IExcelInterface excelInterface)
        {
            _context = context;
            _excelInterface = excelInterface;
        }

        public async Task<IActionResult> Index()
        {
            var produtos = await _context.Produtos.ToListAsync();
            return View(produtos);
        }


        [HttpPost]
        public IActionResult ImportExcel(IFormFile form)
        {
            if (ModelState.IsValid)
            {
                var streamFile = _excelInterface.LerStream(form);
                var produtos = _excelInterface.LerXls(streamFile);
                _excelInterface.SalvarDados(produtos);

                return RedirectToAction("Index");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }
    }
}
