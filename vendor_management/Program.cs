using System.Net;
using SYS_VENDOR_MGMT.Infrastructure;

var builder = WebApplication.CreateBuilder(args);

// Add MVC with Razor Views
builder.Services.AddControllersWithViews();

// Register SOAP sequence client (replaces ASMX Web Reference proxy)
builder.Services.AddScoped<SoapSequenceClient>();

// Session support (replaces System.Web.SessionState)
builder.Services.AddSession(options =>
{
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

// HttpClient factory — used for SOAP calls to Vendor_SEQ / SEQ_service ASMX
builder.Services.AddHttpClient();

// Configuration is available via IConfiguration injection
var app = builder.Build();

// Enforce TLS 1.2 / 1.3 for all outbound ServicePointManager connections
ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseSession();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
