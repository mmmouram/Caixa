using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Moq;
using NUnit.Framework;
using MyApp.Models;
using MyApp.Repositories;
using MyApp.Services;

namespace MyApp.Tests.UnitTests
{
    [TestFixture]
    public class PedidoServiceTests
    {
        private Mock<IPedidoRepository> _pedidoRepositoryMock;
        private IPedidoService _pedidoService;
        private Pedido _pedidoFake;

        [SetUp]
        public void Setup()
        {
            _pedidoRepositoryMock = new Mock<IPedidoRepository>();
            _pedidoService = new PedidoService(_pedidoRepositoryMock.Object);

            _pedidoFake = new Pedido
            {
                Id = 1,
                NumeroPedido = "PED123",
                Cliente = new Cliente { Id = 1, CNPJ = "00.000.000/0001-00", RazaoSocial = "Cliente Teste" },
                PedidoItens = new List<PedidoItem> {
                    new PedidoItem { Id = 1, Descricao = "Item 1", Quantidade = 2, PrecoUnitario = 10.0m },
                    new PedidoItem { Id = 2, Descricao = "Item 2", Quantidade = 1, PrecoUnitario = 20.0m }
                },
                Observacoes = new List<Observacao> {
                    new Observacao { Id = 1, Texto = "Observação 1" }
                },
                Bloqueios = new List<Bloqueio> {
                    new Bloqueio { Id = 1, Motivo = "Bloqueio 1" }
                },
                NotasFiscais = new List<NotaFiscal> {
                    new NotaFiscal { Id = 1, NumeroNota = "NF123", Detalhes = "Detalhes NF" }
                }
            };
        }

        [Test]
        public async Task ObterDetalhesPedidoAsync_WithValidNumeroPedido_ReturnsPedido()
        {
            // Arrange
            _pedidoRepositoryMock.Setup(repo => repo.ObterPedidoPorNumeroAsync(It.IsAny<string>()))
                .ReturnsAsync(_pedidoFake);

            // Act
            var result = await _pedidoService.ObterDetalhesPedidoAsync("PED123");

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual("PED123", result.NumeroPedido);
        }

        [Test]
        public void ObterDetalhesPedidoAsync_WhenPedidoNotFound_ThrowsException()
        {
            // Arrange
            _pedidoRepositoryMock.Setup(repo => repo.ObterPedidoPorNumeroAsync(It.IsAny<string>()))
                .ReturnsAsync((Pedido)null);

            // Act & Assert
            var ex = Assert.ThrowsAsync<Exception>(async () =>
                await _pedidoService.ObterDetalhesPedidoAsync("PED_INEXISTENTE"));
            StringAssert.Contains("não encontrado", ex.Message);
        }

        [Test]
        public async Task ObterDadosAbaAsync_WithValidAba_ReturnsCorrectData()
        {
            // Arrange
            _pedidoRepositoryMock.Setup(repo => repo.ObterPedidoPorNumeroAsync(It.IsAny<string>()))
                .ReturnsAsync(_pedidoFake);

            // Act
            var itensResult = await _pedidoService.ObterDadosAbaAsync("PED123", "itens");
            var obsResult = await _pedidoService.ObterDadosAbaAsync("PED123", "observacoes");
            var bloqueiosResult = await _pedidoService.ObterDadosAbaAsync("PED123", "bloqueios");
            var notasResult = await _pedidoService.ObterDadosAbaAsync("PED123", "notas");

            // Assert
            Assert.AreEqual(2, itensResult.Count);
            Assert.AreEqual(1, obsResult.Count);
            Assert.AreEqual(1, bloqueiosResult.Count);
            Assert.AreEqual(1, notasResult.Count);
        }

        [Test]
        public void ObterDadosAbaAsync_WithInvalidAba_ThrowsException()
        {
            // Arrange
            _pedidoRepositoryMock.Setup(repo => repo.ObterPedidoPorNumeroAsync(It.IsAny<string>()))
                .ReturnsAsync(_pedidoFake);

            // Act & Assert
            var ex = Assert.ThrowsAsync<Exception>(async () =>
                await _pedidoService.ObterDadosAbaAsync("PED123", "invalida"));
            StringAssert.Contains("Aba inválida", ex.Message);
        }

        [Test]
        public async Task ExportarDadosParaExcelAsync_WithDados_ReturnsByteArray()
        {
            // Arrange
            _pedidoRepositoryMock.Setup(repo => repo.ObterPedidoPorNumeroAsync(It.IsAny<string>()))
                .ReturnsAsync(_pedidoFake);
            
            // Act
            var result = await _pedidoService.ExportarDadosParaExcelAsync("PED123", "itens");

            // Assert
            Assert.IsNotNull(result);
            Assert.IsInstanceOf<byte[]>(result);
        }

        [Test]
        public void ExportarDadosParaExcelAsync_WhenNoDados_ThrowsException()
        {
            // Arrange
            var pedidoSemItens = new Pedido
            {
                NumeroPedido = "PED_VAZIO",
                Cliente = new Cliente { CNPJ = "00.000.000/0001-00", RazaoSocial = "Cliente Vazio" },
                PedidoItens = new List<PedidoItem>(),
                Observacoes = new List<Observacao>(),
                Bloqueios = new List<Bloqueio>(),
                NotasFiscais = new List<NotaFiscal>()
            };
            _pedidoRepositoryMock.Setup(repo => repo.ObterPedidoPorNumeroAsync(It.IsAny<string>()))
                .ReturnsAsync(pedidoSemItens);

            // Act & Assert
            var ex = Assert.ThrowsAsync<Exception>(async () =>
                await _pedidoService.ExportarDadosParaExcelAsync("PED_VAZIO", "itens"));
            StringAssert.Contains("Não há dados para exportação", ex.Message);
        }

        [Test]
        public async Task ObterDetalheItemAsync_WithExistingItem_ReturnsItem()
        {
            // Arrange
            _pedidoRepositoryMock.Setup(repo => repo.ObterPedidoPorNumeroAsync(It.IsAny<string>()))
                .ReturnsAsync(_pedidoFake);

            // Act
            var result = await _pedidoService.ObterDetalheItemAsync("PED123", "itens", 1);

            // Assert
            Assert.IsNotNull(result);
            var idValue = (int)result.GetType().GetProperty("Id").GetValue(result, null);
            Assert.AreEqual(1, idValue);
        }

        [Test]
        public void ObterDetalheItemAsync_WithNonExistingItem_ThrowsException()
        {
            // Arrange
            _pedidoRepositoryMock.Setup(repo => repo.ObterPedidoPorNumeroAsync(It.IsAny<string>()))
                .ReturnsAsync(_pedidoFake);

            // Act & Assert
            var ex = Assert.ThrowsAsync<Exception>(async () =>
                await _pedidoService.ObterDetalheItemAsync("PED123", "itens", 999));
            StringAssert.Contains("não encontrado", ex.Message);
        }
    }
}
