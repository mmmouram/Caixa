using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MyApp.Models;
using MyApp.Repositories;
using System.Linq;

namespace MyApp.Services
{
    public class PedidoService : IPedidoService
    {
        private readonly IPedidoRepository _pedidoRepository;

        public PedidoService(IPedidoRepository pedidoRepository)
        {
            _pedidoRepository = pedidoRepository;
        }

        public async Task<Pedido> ObterDetalhesPedidoAsync(string numeroPedido)
        {
            var pedido = await _pedidoRepository.ObterPedidoPorNumeroAsync(numeroPedido);
            if (pedido == null)
                throw new Exception($"Pedido {numeroPedido} não encontrado.");
            return pedido;
        }

        public async Task<List<object>> ObterDadosAbaAsync(string numeroPedido, string aba, string ordenacao = null)
        {
            var pedido = await ObterDetalhesPedidoAsync(numeroPedido);
            List<object> resultado = new List<object>();

            switch (aba.ToLower())
            {
                case "itens":
                    resultado = pedido.PedidoItens.Select(pi => (object)new {
                        pi.Id,
                        pi.Descricao,
                        pi.Quantidade,
                        pi.PrecoUnitario
                    }).ToList();
                    break;
                case "observacoes":
                    resultado = pedido.Observacoes.Select(o => (object)new {
                        o.Id,
                        o.Texto
                    }).ToList();
                    break;
                case "bloqueios":
                    resultado = pedido.Bloqueios.Select(b => (object)new {
                        b.Id,
                        b.Motivo
                    }).ToList();
                    break;
                case "notas":
                    resultado = pedido.NotasFiscais.Select(n => (object)new {
                        n.Id,
                        n.NumeroNota,
                        n.Detalhes
                    }).ToList();
                    break;
                default:
                    throw new Exception("Aba inválida.");
            }

            // Simulação da ordenação (caso necessário) 
            if (!string.IsNullOrEmpty(ordenacao))
            {
                // Ordenação simples por propriedade convertida para string
                resultado = resultado.OrderBy(item => item.GetType().GetProperty(ordenacao)?.GetValue(item, null)).ToList();
            }

            return resultado;
        }

        public async Task<byte[]> ExportarDadosParaExcelAsync(string numeroPedido, string aba)
        {
            // Simulação de exportação para Excel.
            // Em uma implementação real, utilizar uma biblioteca como EPPlus ou ClosedXML para gerar o arquivo Excel.
            var dados = await ObterDadosAbaAsync(numeroPedido, aba);
            if (dados == null || dados.Count == 0)
            {
                throw new Exception("Não há dados para exportação.");
            }

            // Simular a geração de arquivo Excel
            // Retornar um array de bytes vazio para simulação
            return new byte[0];
        }

        public async Task<object> ObterDetalheItemAsync(string numeroPedido, string aba, int itemId)
        {
            var dados = await ObterDadosAbaAsync(numeroPedido, aba);
            var item = dados.FirstOrDefault(d => (int)d.GetType().GetProperty("Id")?.GetValue(d, null) == itemId);
            if (item == null)
                throw new Exception($"Item com ID {itemId} não encontrado na aba {aba}.");
            return item;
        }
    }
}
