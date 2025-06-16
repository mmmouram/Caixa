namespace MyApp.Models
{
    public class NotaFiscal
    {
        public int Id { get; set; }
        public int PedidoId { get; set; }
        public Pedido Pedido { get; set; }
        public string NumeroNota { get; set; }
        public string Detalhes { get; set; }
    }
}