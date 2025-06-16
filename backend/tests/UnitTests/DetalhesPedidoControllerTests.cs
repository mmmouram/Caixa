using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Moq;
using NUnit.Framework;
using MyApp.Controllers;
using MyApp.Models;
using MyApp.Services;

namespace MyApp.Tests.UnitTests
{
    [TestFixture]
    public class DetalhesPedidoControllerTests
    {
        private Mock<IPedidoService> _pedidoServiceMock;
        private DetalhesPedidoController _controller;
        private Pedido _pedidoFake;

        [SetUp]
        public void Setup()
        {
            _pedidoServiceMock = new Mock<IPedidoService>();
            _controller = new DetalhesPedidoController(_pedidoServiceMock.Object);

            _pedidoFake = new Pedido
            {
                Id = 1,
                NumeroPedido = "PED123",
                Cliente = new Cliente { Id = 1, CNPJ = "00.000.000/0001-00", RazaoSocial = "Cliente Teste" },
                PedidoItens = new List<PedidoItem> { new PedidoItem { Id = 1, Descricao = "Item 1", Quantidade = 2, PrecoUnitario = 10.0m } },
                Observacoes = new List<Observacao> { new Observacao { Id = 1, Texto = "Observação 1" } },
                Bloqueios = new List<Bloqueio> { new Bloqueio { Id = 1, Motivo = "Bloqueio 1" } },
                NotasFiscais = new List<NotaFiscal> { new NotaFiscal { Id = 1, NumeroNota = "NF123", Detalhes = "Detalhes NF" } }
            };
        }

        [Test]
        public async Task ObterDetalhesPedido_ReturnsOkResult_WithPedidoDetails()
        {
            // Arrange
            _pedidoServiceMock.Setup(s => s.ObterDetalhesPedidoAsync(It.IsAny<string>()))
                .ReturnsAsync(_pedidoFake);

            // Act
            var result = await _controller.ObterDetalhesPedido("PED123");
            
            // Assert
            Assert.IsInstanceOf<OkObjectResult>(result);
            var okResult = result as OkObjectResult;
            dynamic response = okResult.Value;
            Assert.AreEqual("PED123", response.Pedido);
            Assert.IsNotNull(response.Cliente);
            Assert.IsNotNull(response.Abas);
        }

        [Test]
        public async Task ObterDetalhesPedido_WhenExceptionThrown_ReturnsNotFound()
        {
            // Arrange
            _pedidoServiceMock.Setup(s => s.ObterDetalhesPedidoAsync(It.IsAny<string>()))
                .ThrowsAsync(new Exception("Pedido não encontrado"));

            // Act
            var result = await _controller.ObterDetalhesPedido("PED_INEXISTENTE");

            // Assert
            Assert.IsInstanceOf<NotFoundObjectResult>(result);
            var notFoundResult = result as NotFoundObjectResult;
            dynamic response = notFoundResult.Value;
            StringAssert.Contains("Pedido não encontrado", response.mensagem.ToString());
        }

        [Test]
        public async Task ObterDadosAba_ReturnsOkResult_WithData()
        {
            // Arrange
            var dadosAba = new List<object> { new { Id = 1, Descricao = "Item 1" } };
            _pedidoServiceMock.Setup(s => s.ObterDadosAbaAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()))
                .ReturnsAsync(dadosAba);

            // Act
            var result = await _controller.ObterDadosAba("PED123", "itens");
            
            // Assert
            Assert.IsInstanceOf<OkObjectResult>(result);
            var okResult = result as OkObjectResult;
            Assert.AreEqual(dadosAba, okResult.Value);
        }

        [Test]
        public async Task ObterDadosAba_WhenExceptionThrown_ReturnsBadRequest()
        {
            // Arrange
            _pedidoServiceMock.Setup(s => s.ObterDadosAbaAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()))
                .ThrowsAsync(new Exception("Erro na aba"));

            // Act
            var result = await _controller.ObterDadosAba("PED123", "itens");
            
            // Assert
            Assert.IsInstanceOf<BadRequestObjectResult>(result);
            var badRequest = result as BadRequestObjectResult;
            dynamic response = badRequest.Value;
            StringAssert.Contains("Erro na aba", response.mensagem.ToString());
        }

        [Test]
        public async Task ExportarExcel_ReturnsFileResult_WhenSuccessful()
        {
            // Arrange
            byte[] fileBytes = new byte[] { 1, 2, 3 };
            _pedidoServiceMock.Setup(s => s.ExportarDadosParaExcelAsync(It.IsAny<string>(), It.IsAny<string>()))
                .ReturnsAsync(fileBytes);

            // Act
            var result = await _controller.ExportarExcel("PED123", "itens");
            
            // Assert
            Assert.IsInstanceOf<FileContentResult>(result);
            var fileResult = result as FileContentResult;
            Assert.AreEqual(fileBytes, fileResult.FileContents);
            Assert.AreEqual("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileResult.ContentType);
        }

        [Test]
        public async Task ExportarExcel_WhenExceptionThrown_ReturnsBadRequest()
        {
            // Arrange
            _pedidoServiceMock.Setup(s => s.ExportarDadosParaExcelAsync(It.IsAny<string>(), It.IsAny<string>()))
                .ThrowsAsync(new Exception("Erro na exportação"));

            // Act
            var result = await _controller.ExportarExcel("PED123", "itens");
            
            // Assert
            Assert.IsInstanceOf<BadRequestObjectResult>(result);
            var badRequest = result as BadRequestObjectResult;
            dynamic response = badRequest.Value;
            StringAssert.Contains("Erro na exportação", response.mensagem.ToString());
        }

        [Test]
        public async Task AtualizarDadosAba_ReturnsOkResult_WithUpdatedData()
        {
            // Arrange
            var dadosAba = new List<object> { new { Id = 1, Descricao = "Item Atualizado" } };
            _pedidoServiceMock.Setup(s => s.ObterDadosAbaAsync(It.IsAny<string>(), It.IsAny<string>()))
                .ReturnsAsync(dadosAba);

            // Act
            var result = await _controller.AtualizarDadosAba("PED123", "itens");
            
            // Assert
            Assert.IsInstanceOf<OkObjectResult>(result);
            var okResult = result as OkObjectResult;
            Assert.AreEqual(dadosAba, okResult.Value);
        }

        [Test]
        public async Task ObterDetalheItem_ReturnsOkResult_WithItemDetail()
        {
            // Arrange
            var itemDetalhado = new { Id = 1, Descricao = "Item 1 Detalhado" };
            _pedidoServiceMock.Setup(s => s.ObterDetalheItemAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<int>()))
                .ReturnsAsync(itemDetalhado);

            // Act
            var result = await _controller.ObterDetalheItem("PED123", "itens", 1);
            
            // Assert
            Assert.IsInstanceOf<OkObjectResult>(result);
            var okResult = result as OkObjectResult;
            Assert.AreEqual(itemDetalhado, okResult.Value);
        }

        [Test]
        public async Task ObterDetalheItem_WhenExceptionThrown_ReturnsBadRequest()
        {
            // Arrange
            _pedidoServiceMock.Setup(s => s.ObterDetalheItemAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<int>()))
                .ThrowsAsync(new Exception("Item não encontrado"));

            // Act
            var result = await _controller.ObterDetalheItem("PED123", "itens", 999);
            
            // Assert
            Assert.IsInstanceOf<BadRequestObjectResult>(result);
            var badRequest = result as BadRequestObjectResult;
            dynamic response = badRequest.Value;
            StringAssert.Contains("Item não encontrado", response.mensagem.ToString());
        }
    }
}
