using System.Collections.Generic;

namespace MyApp.Models
{
    public class Pedido
    {
        public int Id { get; set; }
        public string NumeroPedido { get; set; }

        public int ClienteId { get; set; }
        public Cliente Cliente { get; set; }

        public List<PedidoItem> PedidoItens { get; set; }
        public List<Observacao> Observacoes { get; set; }
        public List<Bloqueio> Bloqueios { get; set; }
        public List<NotaFiscal> NotasFiscais { get; set; }
    }
}