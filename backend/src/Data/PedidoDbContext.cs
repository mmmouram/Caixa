using Microsoft.EntityFrameworkCore;
using MyApp.Models;

namespace MyApp.Data
{
    public class PedidoDbContext : DbContext
    {
        public PedidoDbContext(DbContextOptions<PedidoDbContext> options)
            : base(options)
        {
        }

        public DbSet<Pedido> Pedidos { get; set; }
        public DbSet<Cliente> Clientes { get; set; }
        public DbSet<PedidoItem> PedidoItens { get; set; }
        public DbSet<Observacao> Observacoes { get; set; }
        public DbSet<Bloqueio> Bloqueios { get; set; }
        public DbSet<NotaFiscal> NotasFiscais { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            // Configurações adicionais se necessário
        }
    }
}