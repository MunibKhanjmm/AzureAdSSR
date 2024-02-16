using AzureAdSSR.Components;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.Identity.Web.UI;
using Microsoft.Identity.Web;
using Syncfusion.Blazor;
using Microsoft.Graph.ExternalConnectors;
using AzureAdSSR.AccessToken;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.Authorization;
using System.ComponentModel.Design;

namespace AzureAdSSR
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);


            var initialScopes = builder.Configuration["DownstreamApi:Scopes"]?.Split(' ') ?? builder.Configuration["MicrosoftGraph:Scopes"]?.Split(' ');

            // Add services to the container.
            builder.Services.AddRazorComponents()
                .AddInteractiveServerComponents();

            //Register Syncfusion license
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1NAaF1cXmhLYVJ/WmFZfVpgdV9GYVZRR2YuP1ZhSXxXdkdiW39fcHVVRWhUUUY=");

            builder.Services.AddMicrosoftIdentityWebAppAuthentication(builder.Configuration)
            .EnableTokenAcquisitionToCallDownstreamApi(new[] { "Calendars.Read", "Calendars.ReadWrite","User.Read" })
            .AddMicrosoftGraph(builder.Configuration.GetSection("MicrosoftGraph"))
            .AddInMemoryTokenCaches()
              .AddDistributedTokenCaches();

            builder.Services.AddStackExchangeRedisCache(options =>
            {
                options.Configuration = builder.Configuration.GetConnectionString("Redis");
                options.InstanceName = "RedisDemos_"; // unique to the app
            });


            builder.Services.AddControllersWithViews(options =>
            {
                var policy = new AuthorizationPolicyBuilder()
                                 .RequireAuthenticatedUser()
                                 .Build();
                options.Filters.Add(new AuthorizeFilter(policy));
            }).AddMicrosoftIdentityUI();

            builder.Services.AddRazorPages();

            builder.Services.AddAuthorization(options =>
            {
                // By default, all incoming requests will be authorized according to the default policy
                options.FallbackPolicy = options.DefaultPolicy;
            });


            builder.Services.Configure<OpenIdConnectOptions>(OpenIdConnectDefaults.AuthenticationScheme, options =>
            {
                options.Events.OnSignedOutCallbackRedirect = async context =>
                {
                    context.HttpContext.Response.Redirect(context.Options.SignedOutRedirectUri);
                    context.HandleResponse();
                };

            });

            //builder.Services.AddServerSideBlazor();
            builder.Services.AddSyncfusionBlazor();

            builder.Services.AddScoped<MicrosoftIdentityConsentAndConditionalAccessHandler>();

            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (!app.Environment.IsDevelopment())
            {
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();

            app.UseAuthentication();

            app.UseAuthorization();


            app.UseStaticFiles();
            app.UseAntiforgery();

            app.MapControllers();

            app.MapRazorComponents<App>()
                .AddInteractiveServerRenderMode();

            app.Run();
        }
    }
}
