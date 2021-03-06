//------------------------------------------------------------------------------
// Generated by the DevExpress.Blazor package.
// To prevent this operation, add the DxExtendStartupHost property to the project and set this property to False.
//
// BlazorServer_Spreadsheet_MailMerge.csproj:
//
// <Project Sdk="Microsoft.NET.Sdk.Web">
//  <PropertyGroup>
//    <TargetFramework>net5.0</TargetFramework>
//    <DxExtendStartupHost>False</DxExtendStartupHost>
//  </PropertyGroup>
//------------------------------------------------------------------------------
using System;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Hosting;

[assembly: HostingStartup(typeof(BlazorServer_Spreadsheet_MailMerge.DevExpressHostingStartup))]

namespace BlazorServer_Spreadsheet_MailMerge {
    public partial class DevExpressHostingStartup : IHostingStartup {
        void IHostingStartup.Configure(IWebHostBuilder builder) {
            builder.ConfigureServices((serviceCollection) => {
                serviceCollection.AddDevExpressBlazor();
            });
        }
    }
}
