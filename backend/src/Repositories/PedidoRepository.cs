using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using MyApp.Data;
using MyApp.Models;

namespace MyApp.Repositories
{
    public class PedidoRepository : IPedidoRepository
    {
        private readonly PedidoDbContext _context;

        public PedidoRepository(PedidoDbContext context)
        {
            _context = context;
        }

        public async Task<Pedido> ObterPedidoPorNumeroAsync(string numeroPedido)
        {
            // Incluir os relacionamentos com Cliente e demais entidades
            return await _context.Pedidos
                .Include(p => p.Cliente)
                .Include(p => p.PedidoItens)
                .Include(p => p.Observacoes)
                .Include(p => p.Bloqueios)
                .Include(p => p.NotasFiscais)
                .FirstOrDefaultAsync(p => p.NumeroPedido == numeroPedido);
        }
    }
}
