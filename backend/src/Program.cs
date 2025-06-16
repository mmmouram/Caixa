using Microsoft.EntityFrameworkCore;
using Microsoft.OpenApi.Models;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using MyApp.Data;
using MyApp.Config;
using MyApp.Middleware;

var builder = WebApplication.CreateBuilder(args);

// Configurar string de conexão (exemplo, ajustar conforme ambiente)
var connectionString = builder.Configuration.GetConnectionString("PedidoDb")
                       ?? "Server=localhost;Database=PedidoDb;User Id=sa;Password=yourStrong(!)Password;";

// Adicionar serviços ao contêiner
builder.Services.AddControllers();

// Configurar o EF Core para SQL Server
builder.Services.AddDbContext<PedidoDbContext>(options =>
{
    options.UseSqlServer(connectionString);
});

// Configuração de injeção de dependências do projeto
builder.Services.AddInfrastructure();

// Adicionar Swagger para documentação da API
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo { Title = "Pedido API", Version = "v1" });
});

var app = builder.Build();

// Middleware para tratamento global de erros
app.UseMiddleware<ErrorHandlingMiddleware>();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseAuthorization();

app.MapControllers();

app.Run();