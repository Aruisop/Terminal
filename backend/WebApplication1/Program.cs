using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using WebApplication1.Data;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

// Add GraphicWalkerConnection as a scoped service
builder.Services.AddScoped<GraphicWalkerConnection>();

// Add logging
builder.Services.AddLogging();

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll",
        builder =>
        {
            builder.WithOrigins("http://localhost:3000")
                   .AllowAnyMethod()
                   .AllowAnyHeader();
        });
});

builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

// Use the correct policy name that matches what you defined
app.UseCors("AllowAll");  // Changed from "AllowReactApp" to "AllowAll"

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
