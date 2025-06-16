using Microsoft.Extensions.DependencyInjection;
using NUnit.Framework;
using MyApp.Config;
using MyApp.Repositories;
using MyApp.Services;

namespace MyApp.Tests.UnitTests
{
    [TestFixture]
    public class DependencyInjectionTests
    {
        private ServiceCollection _services;

        [SetUp]
        public void SetUp()
        {
            _services = new ServiceCollection();
        }

        [Test]
        public void AddInfrastructure_RegistersRepositoriesAndServices()
        {
            // Act
            _services.AddInfrastructure();
            var serviceProvider = _services.BuildServiceProvider();

            // Assert
            var pedidoRepository = serviceProvider.GetService<IPedidoRepository>();
            var pedidoService = serviceProvider.GetService<IPedidoService>();

            Assert.IsNotNull(pedidoRepository);
            Assert.IsNotNull(pedidoService);
            Assert.IsInstanceOf<PedidoRepository>(pedidoRepository);
            Assert.IsInstanceOf<PedidoService>(pedidoService);
        }
    }
}
