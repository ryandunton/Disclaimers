using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Azure.Core;
using Microsoft.Extensions.Caching.Memory;

var builder = WebApplication.CreateBuilder(args);

// Get configuration values from appsettings.json
var keyVaultUrl = builder.Configuration["KeyVault:Url"] 
    ?? "https://mycompanydisclaimerkv.vault.usgovcloudapi.net/";

// Add Key Vault configuration using managed identity for authentication
builder.Configuration.AddAzureKeyVault(
    new Uri(keyVaultUrl),
    new DefaultAzureCredential(new DefaultAzureCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzureGovernment })
);

// Add services to the container.
builder.Services.AddHttpClient(); // Register IHttpClientFactory
builder.Services.AddControllers(); // Register controllers

// Add Azure Key Vault client as a singleton service, using managed identity for authentication
builder.Services.AddSingleton(provider =>
{
    return new SecretClient(
        new Uri(keyVaultUrl),
        new DefaultAzureCredential(new DefaultAzureCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzureGovernment })
    );
});

// Add Memory Cache
builder.Services.AddMemoryCache();

// Configure CORS for development
if (builder.Environment.IsDevelopment())
{
    builder.Services.AddCors(options =>
    {
        options.AddDefaultPolicy(policy =>
        {
            policy.AllowAnyOrigin()
                  .AllowAnyHeader()
                  .AllowAnyMethod();
        });
    });
}

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}
else
{
    // Enable CORS in development
    app.UseCors();
}

app.UseDefaultFiles();

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

app.UseAuthorization();

app.MapControllers(); // Map attribute-routed controllers

app.Run();
