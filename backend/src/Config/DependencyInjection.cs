using Microsoft.Extensions.DependencyInjection;
using MyApp.Repositories;
using MyApp.Services;

namespace MyApp.Config
{
    public static class DependencyInjection
    {
        public static IServiceCollection AddInfrastructure(this IServiceCollection services)
        {
            // Registrar repositórios
            services.AddScoped<IPedidoRepository, PedidoRepository>();

            // Registrar serviços
            services.AddScoped<IPedidoService, PedidoService>();

            return services;
        }
    }
}