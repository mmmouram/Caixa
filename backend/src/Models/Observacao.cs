namespace MyApp.Models
{
    public class Observacao
    {
        public int Id { get; set; }
        public int PedidoId { get; set; }
        public Pedido Pedido { get; set; }
        public string Texto { get; set; }
    }
}