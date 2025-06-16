using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using MyApp.Services;

namespace MyApp.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DetalhesPedidoController : ControllerBase
    {
        private readonly IPedidoService _pedidoService;

        public DetalhesPedidoController(IPedidoService pedidoService)
        {
            _pedidoService = pedidoService;
        }

        /// <summary>
        /// Retorna os detalhes do pedido, incluindo dados do cliente e informações agregadas.
        /// </summary>
        [HttpGet("{numeroPedido}")]
        public async Task<IActionResult> ObterDetalhesPedido(string numeroPedido)
        {
            try
            {
                var pedido = await _pedidoService.ObterDetalhesPedidoAsync(numeroPedido);

                var response = new
                {
                    Pedido = pedido.NumeroPedido,
                    Cliente = new {
                        pedido.Cliente.CNPJ,
                        pedido.Cliente.RazaoSocial
                    },
                    Abas = new {
                        Itens = pedido.PedidoItens,
                        Observacoes = pedido.Observacoes,
                        Bloqueios = pedido.Bloqueios,
                        NotasFiscais = pedido.NotasFiscais
                    }
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                return NotFound(new { mensagem = ex.Message });
            }
        }

        /// <summary>
        /// Retorna os dados de uma aba específica, com possibilidade de ordenação
        /// </summary>
        [HttpGet("{numeroPedido}/aba/{aba}")]
        public async Task<IActionResult> ObterDadosAba(string numeroPedido, string aba, [FromQuery] string ordenacao = null)
        {
            try
            {
                var dados = await _pedidoService.ObterDadosAbaAsync(numeroPedido, aba, ordenacao);
                return Ok(dados);
            }
            catch (Exception ex)
            {
                return BadRequest(new { mensagem = ex.Message });
            }
        }

        /// <summary>
        /// Exporta os dados da aba para Excel
        /// </summary>
        [HttpPost("{numeroPedido}/aba/{aba}/exportar")]
        public async Task<IActionResult> ExportarExcel(string numeroPedido, string aba)
        {
            try
            {
                var arquivoBytes = await _pedidoService.ExportarDadosParaExcelAsync(numeroPedido, aba);

                // Para fins de exemplo, retornamos um arquivo com nome fictício
                return File(arquivoBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{aba}_pedido_{numeroPedido}.xlsx");
            }
            catch (Exception ex)
            {
                return BadRequest(new { mensagem = ex.Message });
            }
        }

        /// <summary>
        /// Atualiza os dados da aba consultando o banco de dados para dados mais recentes
        /// </summary>
        [HttpPost("{numeroPedido}/aba/{aba}/atualizar")]
        public async Task<IActionResult> AtualizarDadosAba(string numeroPedido, string aba)
        {
            try
            {
                var dados = await _pedidoService.ObterDadosAbaAsync(numeroPedido, aba);
                return Ok(dados);
            }
            catch (Exception ex)
            {
                return BadRequest(new { mensagem = ex.Message });
            }
        }

        /// <summary>
        /// Retorna detalhes adicionais de um item em uma aba ao clicar duas vezes
        /// </summary>
        [HttpGet("{numeroPedido}/aba/{aba}/item/{itemId}")]
        public async Task<IActionResult> ObterDetalheItem(string numeroPedido, string aba, int itemId)
        {
            try
            {
                var itemDetalhado = await _pedidoService.ObterDetalheItemAsync(numeroPedido, aba, itemId);
                return Ok(itemDetalhado);
            }
            catch (Exception ex)
            {
                return BadRequest(new { mensagem = ex.Message });
            }
        }
    }
}
