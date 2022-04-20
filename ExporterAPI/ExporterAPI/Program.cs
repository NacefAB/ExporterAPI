var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseKestrel(options =>
{
    options.Limits.MaxRequestBufferSize = 2147483648;    
    options.Limits.MaxRequestLineSize = 302768;
});

// Add services to the container.

builder.Services.AddControllersWithViews();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddCors();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseCors(builder=>builder.AllowAnyHeader()
                    .AllowAnyOrigin()
                    .AllowAnyMethod());
app.UseAuthorization();

app.MapControllers();

app.Run();
