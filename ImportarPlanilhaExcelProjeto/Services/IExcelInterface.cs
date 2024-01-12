using ImportarPlanilhaExcelProjeto.Models;

namespace ImportarPlanilhaExcelProjeto.Services
{
    public interface IExcelInterface
    {
        MemoryStream LerStream(IFormFile formFile);
        List<ProdutoModel> LerXls(MemoryStream stream);
        void SalvarDados(List<ProdutoModel> produtos);
    }
}
