using System.Threading.Tasks;
using MyApp.Models;

namespace MyApp.Repositories
{
    public interface IPedidoRepository
    {
        Task<Pedido> ObterPedidoPorNumeroAsync(string numeroPedido);
        // Outros métodos para obter dados das abas, se necessário
    }
}
