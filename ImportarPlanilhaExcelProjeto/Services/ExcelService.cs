using ImportarPlanilhaExcelProjeto.Data;
using ImportarPlanilhaExcelProjeto.Models;
using OfficeOpenXml;
using System.ComponentModel;

namespace ImportarPlanilhaExcelProjeto.Services
{
    public class ExcelService : IExcelInterface
    {
        private readonly AppDbContext _context;

        public ExcelService(AppDbContext context)
        {
            _context = context;
        }

        public MemoryStream LerStream(IFormFile formFile)
        {
            using (var stream = new MemoryStream())
            {
                    formFile?.CopyTo(stream);
                    var ListaBytes = stream.ToArray();

                    return new MemoryStream(ListaBytes);
            }
        }

        public List<ProdutoModel> LerXls(MemoryStream stream)
        {
            try
            {
                var resposta = new List<ProdutoModel>();

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int NumeroLinhas = worksheet.Dimension.End.Row;

                    for(int linha = 2; linha <= NumeroLinhas; linha++)
                    {
                        var produto = new ProdutoModel();

                        if (worksheet.Cells[linha, 1].Value != null && worksheet.Cells[linha, 4].Value != null)
                        {
                            produto.Id = 0;
                            produto.Nome = worksheet.Cells[linha, 1].Value.ToString();
                            produto.Valor = Convert.ToDecimal(worksheet.Cells[linha, 2].Value);
                            produto.Quantidade = Convert.ToInt32(worksheet.Cells[linha, 3].Value);
                            produto.Marca = worksheet.Cells[linha, 4].Value.ToString();

                            resposta.Add(produto);
                        }
                        
                    }


                }


                return resposta;

            }catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public  void SalvarDados(List<ProdutoModel> produtos)
        {
            try
            {

                foreach(var produto in produtos) {
                    _context.Add(produto);
                    _context.SaveChanges();           
                }

            }catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
