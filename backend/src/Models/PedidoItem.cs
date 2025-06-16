namespace MyApp.Models
{
    public class PedidoItem
    {
        public int Id { get; set; }
        public int PedidoId { get; set; }
        public Pedido Pedido { get; set; }
        public string Descricao { get; set; }
        public int Quantidade { get; set; }
        public decimal PrecoUnitario { get; set; }
    }
}