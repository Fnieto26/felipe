using Directions_Api;
using Directions_Api.Bussines_Logic;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var configuration = builder.Configuration.Get<Configuration>();

builder.Services.AddScoped<Coordenadas>(x => new Coordenadas(configuration.cadenaConexion));
builder.Services.AddScoped<Archivos>(x => new Archivos(configuration.cadenaConexion, x.GetService<Coordenadas>()));
builder.Services.AddScoped<Seguridad>(x => new Seguridad(configuration.cadenaConexion));
builder.Services.AddScoped<ValidaDireccion>(x => new ValidaDireccion(configuration.cadenaConexion));
builder.Services.AddScoped<Configuration>(x => configuration);

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI();

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
