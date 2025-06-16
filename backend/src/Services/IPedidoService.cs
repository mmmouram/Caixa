using System.Threading.Tasks;
using MyApp.Models;
using System.Collections.Generic;

namespace MyApp.Services
{
    public interface IPedidoService
    {
        Task<Pedido> ObterDetalhesPedidoAsync(string numeroPedido);
        Task<List<object>> ObterDadosAbaAsync(string numeroPedido, string aba, string ordenacao = null);
        Task<byte[]> ExportarDadosParaExcelAsync(string numeroPedido, string aba);
        Task<object> ObterDetalheItemAsync(string numeroPedido, string aba, int itemId);
    }
}
